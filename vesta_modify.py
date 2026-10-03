# -*- coding: utf-8 -*-
"""
VESTA ``.vesta`` 文件内容修改工具
=================================

本模块从 ``mk-ppt/mk_ppt_part_structure`` 项目集成了
``Vesta_File_Modification`` 类，并修复了原实现中的若干 bug。

用途：在 ``CONTCAR/POSCAR -> .vesta`` 之后、导出图片之前，直接修改
``.vesta`` 文件中的显示参数，从而得到统一的作图风格。

支持的段（section）：

================  ==========================================================
段名              作用
================  ==========================================================
``COMPS``         是否显示晶胞坐标轴（``COMPS 1`` / ``COMPS 0``）
``UCOLP``         是否显示晶胞边界线
``SBOND``         化学键（配键）定义，按共价半径和自动生成
``BOUND``         显示/裁剪边界（6 个分数坐标）
``ATOMT``         原子*类型*的半径与颜色
``SITET``         每个*位点*的半径与颜色
``SCENE``         视角旋转矩阵、平移与缩放
================  ==========================================================

基本用法::

    from vesta_modify import VestaFileModification

    vfm = VestaFileModification("CONTCAR.vesta")
    vfm.COMPS_type = "OFF"          # 关闭坐标轴
    vfm.modify_section_COMPS()
    vfm.UCOLP_type = "ON"           # 打开晶胞边界线
    vfm.modify_section_UCOLP()
    vfm.SBOND_type = "ON"           # 生成化学键
    vfm.modify_section_SBOND()
    vfm.BOUNDARY_data = [-0.01, 1.01, -0.01, 1.01, -0.01, 1.01]
    vfm.modify_section_BOUND()
    vfm.modify_atomt([{"atom_name": "Fe", "radius": "1.0", "color_RGB": "255 100 0"}])
    vfm.modify_sitet([...])
    vfm.modify_section_SCENE(0, 0, 1.2, [1, 0, 0, 0, 1, 0, 0, 0, 1])

或者一步到位::

    vfm.apply(
        comps="OFF",
        ucolp="ON",
        sbond="ON",
        boundary=[-0.01, 1.01, -0.01, 1.01, -0.01, 1.01],
        version=[1, 0, 0, 0, 1, 0, 0, 0, 1],
        scale_frac=1.2,
    )

.. note::
   为了兼容原项目，类名 ``Vesta_File_Modification`` 作为别名保留。
"""

from __future__ import annotations

import re
from itertools import combinations_with_replacement
from pathlib import Path

__all__ = [
    "VestaFileModification",
    "Vesta_File_Modification",
    "VestaModifyError",
    "ELEMENTS_RADIUS",
]


class VestaModifyError(RuntimeError):
    """修改 ``.vesta`` 文件失败时抛出。"""


# ---------------------------------------------------------------------------
# 元素共价半径（Å）
# ---------------------------------------------------------------------------
# 数值沿用 mk-ppt 项目（main copy.py）以保持生成键长一致；
# 原表缺失的元素（稀有气体、锕系等）在此补全，未知元素还有默认回退。
ELEMENTS_RADIUS: dict = {
    "H": 0.32, "He": 0.28,
    "Li": 1.23, "Be": 0.89, "B": 0.82, "C": 0.77, "N": 0.70, "O": 0.66,
    "F": 0.64, "Ne": 0.58,
    "Na": 1.54, "Mg": 1.36, "Al": 1.18, "Si": 1.17, "P": 1.10, "S": 1.04,
    "Cl": 0.99, "Ar": 1.06,
    "K": 2.03, "Ca": 1.74, "Sc": 1.44, "Ti": 1.32, "V": 1.22, "Cr": 1.18,
    "Mn": 1.17, "Fe": 1.17, "Co": 1.16, "Ni": 1.15, "Cu": 1.17, "Zn": 1.25,
    "Ga": 1.26, "Ge": 1.22, "As": 1.21, "Se": 1.17, "Br": 1.14, "Kr": 1.16,
    "Rb": 2.16, "Sr": 1.91, "Y": 1.62, "Zr": 1.45, "Nb": 1.34, "Mo": 1.30,
    "Tc": 1.27, "Ru": 1.25, "Rh": 1.25, "Pd": 1.28, "Ag": 1.34, "Cd": 1.48,
    "In": 1.44, "Sn": 1.40, "Sb": 1.41, "Te": 1.37, "I": 1.33, "Xe": 1.40,
    "Cs": 2.35, "Ba": 1.98, "La": 1.87, "Ce": 1.83, "Pr": 1.83, "Nd": 1.82,
    "Pm": 1.80, "Sm": 1.80, "Eu": 2.04, "Gd": 1.80, "Tb": 1.78, "Dy": 1.77,
    "Ho": 1.77, "Er": 1.76, "Tm": 1.75, "Yb": 1.94, "Lu": 1.73, "Hf": 1.44,
    "Ta": 1.34, "W": 1.30, "Re": 1.28, "Os": 1.26, "Ir": 1.27, "Pt": 1.30,
    "Au": 1.34, "Hg": 1.44, "Tl": 1.48, "Pb": 1.47, "Bi": 1.46, "Po": 1.46,
    "At": 1.45, "Rn": 1.50,
    "Fr": 2.60, "Ra": 2.21, "Ac": 2.15, "Th": 2.06, "Pa": 2.00, "U": 1.45,
    "Np": 1.90, "Pu": 1.87, "Am": 1.80, "Cm": 1.69,
}

DEFAULT_RADIUS = 1.40          # 未收录元素的回退共价半径
DEFAULT_SBOND_PADDING = 0.50   # 成键判据：r_a + r_b + padding


class VestaFileModification:
    """读取并原地修改一个 ``.vesta`` 文件。

    Parameters
    ----------
    vesta_file_path:
        ``.vesta`` 文件路径。所有修改都会写回该文件。
    sbond_padding:
        自动成键的额外容差（Å），键长上限 = ``r_a + r_b + sbond_padding``。
    default_radius:
        当元素不在 :data:`ELEMENTS_RADIUS` 中时使用的回退半径。
    verbose:
        是否输出提示信息。
    """

    def __init__(
        self,
        vesta_file_path: str | Path,
        sbond_padding: float = DEFAULT_SBOND_PADDING,
        default_radius: float = DEFAULT_RADIUS,
        verbose: bool = True,
    ) -> None:
        self.file_path = Path(vesta_file_path)
        if not self.file_path.exists():
            raise FileNotFoundError(f"找不到 vesta 文件: {self.file_path}")

        self.sbond_padding = float(sbond_padding)
        self.default_radius = float(default_radius)
        self.verbose = verbose

        self.data = self.read_file()
        self.section_lst = self.get_section_lst()

        # 兼容原项目：保留实例级别的元素半径表引用
        self.elements_radius_data = ELEMENTS_RADIUS

        # 从 STRUC 段解析元素种类，再枚举所有元素对计算成键键长
        self.chem_label = self._parse_chem_labels()
        self.combinations_list = list(combinations_with_replacement(self.chem_label, 2))
        self.sbond_length_lst: list = []
        for atom_a, atom_b in self.combinations_list:
            max_len = round(self._radius(atom_a) + self._radius(atom_b) + self.sbond_padding, 2)
            self.sbond_length_lst.append((atom_a, atom_b, max_len))

        # 默认开关
        self.COMPS_type = "ON"
        self.UCOLP_type = "ON"
        self.SBOND_type = "ON"
        self.BOUNDARY_data = [0, 1, 0, 1, 0, 1]

    # ------------------------------------------------------------------ #
    # 日志 / 内部工具
    # ------------------------------------------------------------------ #
    def _log(self, msg: str) -> None:
        if self.verbose:
            print(f"[VestaModify] {msg}", flush=True)

    def _radius(self, symbol: str) -> float:
        """查询元素共价半径，未收录时回退并提示。"""
        radius = ELEMENTS_RADIUS.get(symbol)
        if radius is None:
            self._log(f"未收录元素 {symbol!r} 的共价半径，使用默认值 {self.default_radius} A")
            return self.default_radius
        return radius

    # ------------------------------------------------------------------ #
    # 读写
    # ------------------------------------------------------------------ #
    def read_file(self) -> str:
        with open(self.file_path, "r", encoding="utf-8") as f:
            return f.read()

    def write_file(self) -> None:
        with open(self.file_path, "w", encoding="utf-8") as f:
            f.write(self.data)

    # ------------------------------------------------------------------ #
    # 解析
    # ------------------------------------------------------------------ #
    def get_section_lst(self) -> list:
        """返回文件中所有“以字母开头的行”的首个关键字。"""
        section_lst = []
        for line in self.data.splitlines():
            stripped = line.strip()
            if stripped and stripped[0].isalpha():
                section_lst.append(stripped.split()[0])
        return section_lst

    @staticmethod
    def _first_token(line: str) -> str:
        stripped = line.strip()
        return stripped.split()[0] if stripped else ""

    def extract_section_msg(self, section_label: str) -> list:
        """提取 ``section_label`` 段的内容（不含段标题行）。

        - ``TITLE``：返回标题行的 ``split()`` 结果（标题在关键字下一行）；
        - ``COMPS``：关键字与值在同一行，返回该行 ``split()``；
        - 其它：返回关键字之后的每一行（``rstrip`` 后），直到下一个字母开头的行。

        找不到时返回空列表（原实现会返回 ``None``，导致后续 ``.split()`` 报错）。
        """
        lines = self.data.splitlines()

        if section_label == "TITLE":
            for i, line in enumerate(lines):
                if self._first_token(line) == "TITLE" and i + 1 < len(lines):
                    return lines[i + 1].split()
            return []

        if section_label == "COMPS":
            for line in lines:
                if self._first_token(line) == "COMPS":
                    return line.split()
            return []

        section_msg: list = []
        in_section = False
        for line in lines:
            stripped = line.strip()
            if stripped and stripped[0].isalpha():
                if in_section:
                    break
                if stripped.split()[0] == section_label:
                    in_section = True
                    continue
            if in_section:
                section_msg.append(line.rstrip())
        return section_msg

    def _parse_chem_labels(self) -> list:
        """从 STRUC 段提取元素符号（按出现顺序、去重）。

        通过“首列是纯数字且第二列是元素符号”识别原子头行，
        比原实现的“奇偶行”判断更稳健。
        """
        labels: list = []
        for line in self.extract_section_msg("STRUC"):
            parts = line.split()
            if len(parts) < 2:
                continue
            # 原子头行形如：  1  H         H1  1.0000 ...
            if parts[0].isdigit() and parts[0] != "0" and parts[1] and parts[1][0].isalpha():
                label = parts[1]
                if label not in labels:
                    labels.append(label)
        return labels

    # ------------------------------------------------------------------ #
    # 重写整段
    # ------------------------------------------------------------------ #
    def re_write_section_msg(self, section_label: str, section_msg: str) -> None:
        """用 ``section_msg`` 替换文件中的 ``section_label`` 段。

        原实现用 ``section_label in line`` 做子串匹配，可能误命中；
        这里改为“首个关键字完全相等”，并且找不到段时抛出明确异常
        （原实现会把新段错误地追加到文件末尾）。
        """
        lines = self.data.splitlines(keepends=True)
        part_a: list = []
        part_b: list = []
        found = False
        skipping = False

        for line in lines:
            stripped = line.strip()
            first = stripped.split()[0] if stripped else ""

            if not found:
                if first == section_label:
                    found = True
                    skipping = True
                    continue
                part_a.append(line)
                continue

            if skipping:
                # 段内容是非字母开头的行；遇到下一个字母行即结束
                if stripped and stripped[0].isalpha():
                    skipping = False
                    part_b.append(line)
                # 否则继续跳过旧段内容
            else:
                part_b.append(line)

        if not found:
            raise VestaModifyError(f"在 {self.file_path} 中找不到 {section_label} 段")

        self.data = "".join(part_a) + section_msg + "\n" + "".join(part_b)
        self.write_file()

    # ------------------------------------------------------------------ #
    # 各段修改
    # ------------------------------------------------------------------ #
    def modify_section_COMPS(self) -> None:
        """设置是否显示晶胞坐标轴（``COMPS 1`` 开 / ``COMPS 0`` 关）。"""
        value = "1" if str(self.COMPS_type).upper() == "ON" else "0"
        new_data, n = re.subn(r"(?m)^(COMPS[ \t]+)\S+", lambda m: m.group(1) + value, self.data)
        if not n:
            self._log("未找到 COMPS 段，已跳过")
            return
        self.data = new_data
        self.write_file()

    def modify_section_UCOLP(self) -> None:
        """设置是否显示晶胞边界线（UCOLP 数据行的第 2 个数值，1 开 / 0 关）。"""
        value = "1" if str(self.UCOLP_type).upper() == "ON" else "0"
        # UCOLP 段首行是关键字，下一行是数据行：0   <flag>   1.000 ...
        pattern = re.compile(
            r"(?m)(^UCOLP[ \t]*\n[ \t]*[-+]?\d+(?:\.\d+)?[ \t]+)"  # UCOLP + 第 1 个数值
            r"([-+]?\d+(?:\.\d+)?)"                                  # 第 2 个数值（要改的 flag）
        )
        new_data, n = pattern.subn(lambda m: m.group(1) + value, self.data)
        if not n:
            self._log("未找到 UCOLP 数据行，已跳过")
            return
        self.data = new_data
        self.write_file()

    def modify_section_BOUND(self) -> None:
        """设置显示边界。``BOUNDARY_data`` 为 6 个分数坐标。"""
        values = "      ".join(str(v) for v in self.BOUNDARY_data)
        section_msg = f"BOUND\n   {values}\n  0   0   0   0  0"
        self.re_write_section_msg("BOUND", section_msg)

    def modify_section_SBOND(self) -> None:
        """按元素半径和重建 SBOND 段。

        修复：原实现 ``atom_atom_line_num`` 从未自增，导致所有键的序号都是 1。

        ``SBOND_type`` 为 ``OFF`` 时保持 VESTA 自动生成的键定义不变。
        """
        if str(self.SBOND_type).upper() != "ON":
            return

        lines = ["SBOND"]
        for idx, (atom_a, atom_b, max_len) in enumerate(self.sbond_length_lst, start=1):
            lines.append(
                f"{idx}   {atom_a}    {atom_b}    0.00000   {max_len}  "
                "0  0  1  0  1  0.250  2.000 127 127 127"
            )
        lines.append("  0 0 0 0")
        self.re_write_section_msg("SBOND", "\n".join(lines))

    def get_scene_matrix(self) -> list | None:
        """读取当前 SCENE 段的 3x3 视角矩阵（9 个数字），失败返回 ``None``。"""
        vals: list = []
        for line in self.extract_section_msg("SCENE")[:3]:
            parts = line.split()
            if len(parts) < 3:
                return None
            try:
                vals.extend(float(x) for x in parts[:3])
            except ValueError:
                return None
        return vals if len(vals) == 9 else None

    def modify_section_SCENE(self, x_move_frac, y_move_frac, scal_frac, vesion_data) -> None:
        """设置视角。

        Parameters
        ----------
        vesion_data:
            9 个数字，按 ``ax, ay, az, bx, by, bz, cx, cy, cz`` 顺序的视角矩阵。
        """
        vals = list(vesion_data)
        if len(vals) != 9:
            raise ValueError(f"vesion_data 需要 9 个数字（3x3 视角矩阵），收到 {len(vals)} 个")
        ax, ay, az, bx, by, bz, cx, cy, cz = vals

        section_msg = (
            "SCENE\n"
            f"{ax}  {ay}   {az}   0.000000\n"
            f"{bx}  {by}   {bz}   0.000000\n"
            f"{cx}  {cy}   {cz}   0.000000\n"
            "0.000000  0.000000  0.000000  1.000000\n"
            f"{x_move_frac}   {y_move_frac}  \n"
            "0.000\n"
            f"{scal_frac}"
        )
        self.re_write_section_msg("SCENE", section_msg)

    # ------------------------------------------------------------------ #
    # 原子半径 / 颜色
    # ------------------------------------------------------------------ #
    @staticmethod
    def _find_params(label: str, atom_parms) -> dict | None:
        for parm in atom_parms or []:
            if parm.get("atom_name") == label:
                return parm
        return None

    @staticmethod
    def _apply_params_to_parts(parts: list, parm: dict) -> None:
        if "radius" in parm and parm["radius"] is not None:
            parts[2] = str(parm["radius"])
        rgb = str(parm.get("color_RGB", "")).split()
        if len(rgb) >= 3:
            parts[3], parts[4], parts[5] = rgb[0], rgb[1], rgb[2]

    def modify_atomt(self, atom_parms) -> None:
        """修改 ATOM T 段（按原子类型设置半径与颜色）。

        ``atom_parms`` 形如
        ``[{"atom_name": "H", "radius": "0.46", "color_RGB": "255 204 204"}, ...]``
        """
        if not atom_parms:
            return

        section_msg = "ATOMT\n"
        for line in self.extract_section_msg("ATOMT"):
            parts = line.split()
            if len(parts) >= 6:
                parm = self._find_params(parts[1], atom_parms)
                if parm:
                    self._apply_params_to_parts(parts, parm)
            section_msg += "  ".join(parts) + "\n"
        # 去掉末尾多余换行，由 re_write_section_msg 统一补一个
        self.re_write_section_msg("ATOMT", section_msg[:-1])

    def modify_sitet(self, atom_parms) -> None:
        """修改 SITET 段（按位点设置半径与颜色，标签会去掉尾部数字）。"""
        if not atom_parms:
            return

        section_msg = "SITET\n"
        for line in self.extract_section_msg("SITET"):
            parts = line.split()
            if len(parts) >= 6:
                label = re.sub(r"\d+", "", parts[1])
                parm = self._find_params(label, atom_parms)
                if parm:
                    self._apply_params_to_parts(parts, parm)
            section_msg += "  ".join(parts) + "\n"
        self.re_write_section_msg("SITET", section_msg[:-1])

    # ------------------------------------------------------------------ #
    # 一步到位
    # ------------------------------------------------------------------ #
    def apply(
        self,
        comps: str | None = None,
        ucolp: str | None = None,
        sbond: str | None = None,
        boundary=None,
        atom_params=None,
        version=None,
        x_move_frac=0.0,
        y_move_frac=0.0,
        scale_frac=1.0,
    ) -> "VestaFileModification":
        """按给定参数批量应用修改（``None`` 表示保持原样）。

        Returns
        -------
        VestaFileModification
            返回 ``self``，便于链式调用。
        """
        if comps is not None:
            self.COMPS_type = comps
            self.modify_section_COMPS()
        if ucolp is not None:
            self.UCOLP_type = ucolp
            self.modify_section_UCOLP()
        if sbond is not None:
            self.SBOND_type = sbond
            self.modify_section_SBOND()
        if boundary is not None:
            self.BOUNDARY_data = list(boundary)
            self.modify_section_BOUND()
        if atom_params:
            self.modify_atomt(atom_params)
            self.modify_sitet(atom_params)
        if version is not None:
            self.modify_section_SCENE(x_move_frac, y_move_frac, scale_frac, version)
        elif x_move_frac or y_move_frac or scale_frac != 1.0:
            # 未显式给出视角矩阵，但需要改平移/缩放：沿用文件里现有的矩阵
            existing = self.get_scene_matrix()
            if existing is None:
                raise VestaModifyError("无法从 SCENE 段读取现有视角矩阵，请显式提供 version")
            self.modify_section_SCENE(x_move_frac, y_move_frac, scale_frac, existing)
        return self


# 兼容原项目的类名
Vesta_File_Modification = VestaFileModification
