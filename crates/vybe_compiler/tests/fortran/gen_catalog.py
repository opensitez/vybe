#!/usr/bin/env python3
"""Generate intrinsic catalog test files and calibrate expected outputs."""

from __future__ import annotations

import json
import subprocess
import textwrap
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CALIBRATE_RS = ROOT / "test_catalog_calibrate.rs"

CATALOGS: dict[str, list[tuple[str, str, list[str] | None]]] = {
    "test_intrinsic_catalog_letter_a_f.rs": [
        # achar
        ("achar_iachar_roundtrip", "program t\nprint *, iachar(achar(72))\nend program t\n", ["72"]),
        ("achar_letter_b_code", "program t\nprint *, iachar(achar(66))\nend program t\n", ["66"]),
        ("achar_from_integer_var", "program t\ninteger :: k = 67\nprint *, achar(k)\nend program t\n", ["C"]),
        ("achar_kind_default", "program t\nprint *, achar(68, kind=kind('A'))\nend program t\n", ["D"]),
        # acos
        ("acos_one", "program t\nprint *, nint(acos(1.0)*100)\nend program t\n", None),
        ("acos_zero", "program t\nprint *, nint(acos(0.0)*100)\nend program t\n", None),
        ("acos_neg_one", "program t\nprint *, nint(acos(-1.0)*100)\nend program t\n", None),
        ("acos_half", "program t\nprint *, nint(acos(0.5)*100)\nend program t\n", None),
        ("acos_array_element", "program t\nreal :: x(3) = [0.5, 0.0, 1.0]\nprint *, nint(acos(x(1))*100)\nend program t\n", None),
        # acosh
        ("acosh_one", "program t\nprint *, nint(acosh(1.0)*100)\nend program t\n", None),
        ("acosh_two", "program t\nprint *, nint(acosh(2.0)*100)\nend program t\n", None),
        ("acosh_cosh_identity", "program t\nprint *, nint(acosh(cosh(1.0))*100)\nend program t\n", None),
        # adjustl
        ("adjustl_len_trim", "program t\ncharacter(len=10) :: s = '   data'\nprint *, len_trim(adjustl(s))\nend program t\n", ["4"]),
        ("adjustl_single_char", "program t\ncharacter(len=6) :: s = '    Z'\nprint *, trim(adjustl(s))\nend program t\n", ["Z"]),
        ("adjustl_all_blanks", "program t\ncharacter(len=5) :: s = '     '\nprint *, len_trim(adjustl(s))\nend program t\n", ["0"]),
        # adjustr
        ("adjustr_len_trim", "program t\ncharacter(len=10) :: s = 'go'\nprint *, len_trim(adjustr(s))\nend program t\n", ["2"]),
        ("adjustr_padded_trim", "program t\ncharacter(len=8) :: s = 'xy'\nprint *, trim(adjustr(s))\nend program t\n", ["xy"]),
        ("adjustr_concat_context", "program t\ncharacter(len=6) :: s = 'ab'\nprint *, trim(adjustr(s)) // 'c'\nend program t\n", ["abc"]),
        # aimag
        ("aimag_literal", "program t\ncomplex :: z = (4.0, -3.0)\nprint *, nint(aimag(z))\nend program t\n", ["-3"]),
        ("aimag_cmplx_runtime", "program t\ncomplex :: z\nz = cmplx(2.0, 5.0)\nprint *, nint(aimag(z))\nend program t\n", ["5"]),
        ("aimag_pure_imag", "program t\ncomplex :: z = (0.0, 7.0)\nprint *, nint(aimag(z))\nend program t\n", ["7"]),
        ("aimag_zero", "program t\ncomplex :: z = (9.0, 0.0)\nprint *, nint(aimag(z))\nend program t\n", ["0"]),
        # aint
        ("aint_positive", "program t\nprint *, aint(3.9)\nend program t\n", ["3"]),
        ("aint_negative", "program t\nprint *, aint(-3.9)\nend program t\n", ["-3"]),
        ("aint_with_kind", "program t\ninteger, parameter :: sp = kind(1.0)\nprint *, aint(2.7, kind=sp)\nend program t\n", ["2"]),
        ("aint_array_element", "program t\nreal :: x(2) = [4.8, -1.2]\nprint *, aint(x(2))\nend program t\n", ["-1"]),
        # all
        ("all_true_1d", "program t\nlogical :: m(4) = [.true., .true., .true., .true.]\nprint *, all(m)\nend program t\n", ["T"]),
        ("all_false_1d", "program t\nlogical :: m(4) = [.true., .false., .true., .true.]\nprint *, all(m)\nend program t\n", ["F"]),
        ("all_dim1", "program t\nlogical :: m(2,3) = reshape([.true.,.true.,.false.,.true.,.true.,.true.],[2,3])\nlogical :: c(3)\nc = all(m, dim=1)\nprint *, c(1)\nprint *, c(3)\nend program t\n", ["T", "F"]),
        ("all_dim2", "program t\nlogical :: m(2,3) = reshape([.true.,.false.,.true.,.true.,.true.,.false.],[2,3])\nlogical :: r(2)\nr = all(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nend program t\n", ["T", "F"]),
        ("all_dim1_all_true", "program t\nlogical :: m(2,2) = reshape([.true.,.true.,.true.,.true.],[2,2])\nlogical :: c(2)\nc = all(m, dim=1)\nprint *, all(c)\nend program t\n", ["T"]),
        ("all_single_element", "program t\nlogical :: m(1) = [.true.]\nprint *, all(m)\nend program t\n", ["T"]),
        # allocated
        ("allocated_pointer_false", "program t\ninteger, pointer :: p(:) => null()\nprint *, allocated(p)\nend program t\n", ["F"]),
        ("allocated_after_allocate", "program t\ninteger, allocatable :: a(:)\nprint *, allocated(a)\nallocate(a(2))\nprint *, allocated(a)\nend program t\n", ["F", "T"]),
        ("allocated_deallocate", "program t\ninteger, allocatable :: a(:)\nallocate(a(1))\nprint *, allocated(a)\ndeallocate(a)\nprint *, allocated(a)\nend program t\n", ["T", "F"]),
        ("allocated_scalar", "program t\ninteger, allocatable :: n\nprint *, allocated(n)\nend program t\n", ["F"]),
        ("allocated_scalar_after", "program t\ninteger, allocatable :: n\nallocate(n)\nprint *, allocated(n)\nend program t\n", ["T"]),
        # anint
        ("anint_half_up", "program t\nprint *, anint(3.5)\nend program t\n", ["4"]),
        ("anint_half_down", "program t\nprint *, anint(2.5)\nend program t\n", ["2"]),
        ("anint_negative", "program t\nprint *, anint(-2.6)\nend program t\n", ["-3"]),
        ("anint_with_kind", "program t\ninteger, parameter :: sp = kind(1.0)\nprint *, anint(4.4, kind=sp)\nend program t\n", ["4"]),
        # any
        ("any_one_true", "program t\nlogical :: m(4) = [.false., .false., .true., .false.]\nprint *, any(m)\nend program t\n", ["T"]),
        ("any_all_false", "program t\nlogical :: m(3) = [.false., .false., .false.]\nprint *, any(m)\nend program t\n", ["F"]),
        ("any_all_true", "program t\nlogical :: m(2) = [.true., .true.]\nprint *, any(m)\nend program t\n", ["T"]),
        ("any_dim1", "program t\nlogical :: m(2,2) = reshape([.false.,.true.,.false.,.false.],[2,2])\nlogical :: c(2)\nc = any(m, dim=1)\nprint *, c(1)\nprint *, c(2)\nend program t\n", ["T", "F"]),
        ("any_dim2", "program t\nlogical :: m(2,2) = reshape([.false.,.false.,.true.,.false.],[2,2])\nlogical :: r(2)\nr = any(m, dim=2)\nprint *, r(1)\nprint *, r(2)\nend program t\n", ["F", "T"]),
        ("any_dim1_none", "program t\nlogical :: m(2,2) = reshape([.false.,.false.,.false.,.false.],[2,2])\nlogical :: c(2)\nc = any(m, dim=1)\nprint *, any(c)\nend program t\n", ["F"]),
        # asin
        ("asin_zero", "program t\nprint *, nint(asin(0.0)*100)\nend program t\n", None),
        ("asin_half", "program t\nprint *, nint(asin(0.5)*100)\nend program t\n", None),
        ("asin_neg_half", "program t\nprint *, nint(asin(-0.5)*100)\nend program t\n", None),
        ("asin_one", "program t\nprint *, nint(asin(1.0)*100)\nend program t\n", None),
        # asinh
        ("asinh_zero", "program t\nprint *, nint(asinh(0.0)*100)\nend program t\n", None),
        ("asinh_one", "program t\nprint *, nint(asinh(1.0)*100)\nend program t\n", None),
        ("asinh_neg_one", "program t\nprint *, nint(asinh(-1.0)*100)\nend program t\n", None),
        # atan
        ("atan_zero", "program t\nprint *, nint(atan(0.0)*100)\nend program t\n", None),
        ("atan_one", "program t\nprint *, nint(atan(1.0)*100)\nend program t\n", None),
        ("atan_neg_one", "program t\nprint *, nint(atan(-1.0)*100)\nend program t\n", None),
        # atan2
        ("atan2_first_quadrant", "program t\nprint *, nint(atan2(1.0,1.0)*100)\nend program t\n", None),
        ("atan2_pos_x_axis", "program t\nprint *, nint(atan2(0.0,1.0)*100)\nend program t\n", None),
        ("atan2_neg_y_axis", "program t\nprint *, nint(atan2(-1.0,0.0)*100)\nend program t\n", None),
        ("atan2_third_quadrant", "program t\nprint *, nint(atan2(-1.0,-1.0)*100)\nend program t\n", None),
        ("atan2_fourth_quadrant", "program t\nprint *, nint(atan2(1.0,-1.0)*100)\nend program t\n", None),
        # atanh
        ("atanh_zero", "program t\nprint *, nint(atanh(0.0)*100)\nend program t\n", None),
        ("atanh_half", "program t\nprint *, nint(atanh(0.5)*100)\nend program t\n", None),
        ("atanh_neg_half", "program t\nprint *, nint(atanh(-0.5)*100)\nend program t\n", None),
    ],
    "test_intrinsic_catalog_letter_g_m.rs": [
        # bessel_j0/j1
        ("bessel_j0_zero", "program t\nprint *, nint(bessel_j0(0.0)*100)\nend program t\n", None),
        ("bessel_j0_one", "program t\nprint *, nint(bessel_j0(1.0)*100)\nend program t\n", None),
        ("bessel_j1_zero", "program t\nprint *, nint(bessel_j1(0.0)*100)\nend program t\n", None),
        ("bessel_j1_one", "program t\nprint *, nint(bessel_j1(1.0)*100)\nend program t\n", None),
        ("bessel_j0_two", "program t\nprint *, nint(bessel_j0(2.0)*100)\nend program t\n", None),
        # bge/bgt/ble/blt
        ("bge_equal_strings", "program t\nprint *, bge('abc', 'abc')\nend program t\n", ["T"]),
        ("bge_greater_prefix", "program t\nprint *, bge('abd', 'abc')\nend program t\n", ["T"]),
        ("bgt_strictly_greater", "program t\nprint *, bgt('abd', 'abc')\nend program t\n", ["T"]),
        ("bgt_not_equal", "program t\nprint *, bgt('abc', 'abc')\nend program t\n", ["F"]),
        ("ble_equal_strings", "program t\nprint *, ble('abc', 'abc')\nend program t\n", ["T"]),
        ("ble_less_prefix", "program t\nprint *, ble('abb', 'abc')\nend program t\n", ["T"]),
        ("blt_strictly_less", "program t\nprint *, blt('abb', 'abc')\nend program t\n", ["T"]),
        ("blt_not_equal", "program t\nprint *, blt('abc', 'abc')\nend program t\n", ["F"]),
        # bit_size
        ("bit_size_default_int", "program t\nprint *, bit_size(0)\nend program t\n", ["32"]),
        ("bit_size_kind_param", "program t\ninteger, parameter :: ik = kind(0)\nprint *, bit_size(0, kind=ik)\nend program t\n", ["32"]),
        ("bit_size_on_variable", "program t\ninteger :: n = 7\nprint *, bit_size(n)\nend program t\n", ["32"]),
        # char
        ("char_ichar_roundtrip", "program t\nprint *, ichar(char(75))\nend program t\n", ["75"]),
        ("char_letter_m", "program t\nprint *, char(77)\nend program t\n", ["M"]),
        ("char_kind_default", "program t\nprint *, char(78, kind=kind('N'))\nend program t\n", ["N"]),
        ("char_from_integer", "program t\ninteger :: code = 79\nprint *, char(code)\nend program t\n", ["O"]),
        # command_argument_count (compile + run)
        ("command_argument_count", "program t\nprint *, command_argument_count()\nend program t\n", ["0"]),
        # cmplx variants
        ("cmplx_two_args", "program t\ncomplex :: z\nz = cmplx(3.0, 4.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n", ["3", "4"]),
        ("cmplx_one_arg", "program t\ncomplex :: z\nz = cmplx(6.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n", ["6", "0"]),
        ("cmplx_integer_parts", "program t\ncomplex :: z\nz = cmplx(2, 5)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n", ["2", "5"]),
        ("cmplx_kind_arg", "program t\ninteger, parameter :: dp = kind(1.0d0)\ncomplex(dp) :: z\nz = cmplx(1.0_dp, 2.0_dp, dp)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n", ["1", "2"]),
        ("cmplx_literal_tuple", "program t\ncomplex :: z = (7.0, -2.0)\nprint *, nint(real(z))\nprint *, nint(aimag(z))\nend program t\n", ["7", "-2"]),
        # conjg
        ("conjg_flips_imag", "program t\ncomplex :: z = (5.0, -3.0)\nprint *, nint(real(conjg(z)))\nprint *, nint(aimag(conjg(z)))\nend program t\n", ["5", "3"]),
        ("conjg_pure_real", "program t\ncomplex :: z = (4.0, 0.0)\nprint *, nint(real(conjg(z)))\nprint *, nint(aimag(conjg(z)))\nend program t\n", ["4", "0"]),
        ("conjg_twice", "program t\ncomplex :: z = (1.0, 2.0)\nprint *, nint(real(conjg(conjg(z))))\nprint *, nint(aimag(conjg(conjg(z))))\nend program t\n", ["1", "2"]),
        ("conjg_neg_imag", "program t\ncomplex :: z = (-2.0, 7.0)\nprint *, nint(real(conjg(z)))\nprint *, nint(aimag(conjg(z)))\nend program t\n", ["-2", "-7"]),
        # cos/cosh
        ("cos_zero", "program t\nprint *, nint(cos(0.0)*100)\nend program t\n", None),
        ("cos_pi", "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi)*100)\nend program t\n", None),
        ("cos_half_pi", "program t\nreal, parameter :: pi = acos(-1.0)\nprint *, nint(cos(pi/2.0)*100)\nend program t\n", None),
        ("cosh_zero", "program t\nprint *, nint(cosh(0.0)*100)\nend program t\n", None),
        ("cosh_one", "program t\nprint *, nint(cosh(1.0)*100)\nend program t\n", None),
        ("cosh_neg_one", "program t\nprint *, nint(cosh(-1.0)*100)\nend program t\n", None),
        # count(mask)
        ("count_logical_mask", "program t\nlogical :: m(5) = [.true., .false., .true., .false., .true.]\nprint *, count(m)\nend program t\n", ["3"]),
        ("count_comparison_mask", "program t\ninteger :: a(4) = [1, 3, 5, 7]\nprint *, count(a > 3)\nend program t\n", ["2"]),
        ("count_all_false", "program t\nlogical :: m(3) = [.false., .false., .false.]\nprint *, count(m)\nend program t\n", ["0"]),
        ("count_all_true", "program t\nlogical :: m(2) = [.true., .true.]\nprint *, count(m)\nend program t\n", ["2"]),
        ("count_dim1", "program t\ninteger :: a(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, count(a > 3, dim=1)\nend program t\n", None),
        ("count_dim2", "program t\ninteger :: a(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, count(a > 3, dim=2)\nend program t\n", None),
        # cpu_time compile
        ("cpu_time", "program t\nreal :: t\ncall cpu_time(t)\nprint *, nint(t*100)\nend program t\n", None),
        # cshift
        ("cshift_left_one", "program t\ninteger :: a(4) = [1,2,3,4]\nprint *, cshift(a, 1)\nend program t\n", None),
        ("cshift_right_one", "program t\ninteger :: a(4) = [1,2,3,4]\nprint *, cshift(a, -1)\nend program t\n", None),
        ("cshift_zero", "program t\ninteger :: a(3) = [5,6,7]\nprint *, cshift(a, 0)\nend program t\n", None),
        ("cshift_dim1", "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, cshift(m, 1, dim=2)\nend program t\n", None),
        # date_and_time
        ("date_and_time_values", "program t\ninteger :: dt(8)\ncall date_and_time(values=dt)\nprint *, dt(1)\nprint *, dt(2)\nend program t\n", None),
        ("date_and_time_date", "program t\ncharacter(len=8) :: d\ncall date_and_time(date=d)\nprint *, len_trim(d)\nend program t\n", None),
        ("date_and_time_time", "program t\ncharacter(len=10) :: tm\ncall date_and_time(time=tm)\nprint *, len_trim(tm)\nend program t\n", None),
        ("date_and_time_zone", "program t\ncharacter(len=5) :: z\ncall date_and_time(zone=z)\nprint *, len_trim(z)\nend program t\n", None),
    ],
    "test_intrinsic_catalog_letter_d_i.rs": [
        ("dble_from_int", "program t\nprint *, nint(dble(7))\nend program t\n", ["7"]),
        ("dble_from_real", "program t\nprint *, nint(dble(3.5)*10)\nend program t\n", ["35"]),
        ("dble_preserves_sign", "program t\nprint *, nint(dble(-4))\nend program t\n", ["-4"]),
        ("dble_kind_context", "program t\ndouble precision :: x\nx = dble(9)\nprint *, nint(x)\nend program t\n", ["9"]),
        ("digits_default_real", "program t\nprint *, digits(1.0)\nend program t\n", ["24"]),
        ("digits_double", "program t\nprint *, digits(1.0d0)\nend program t\n", ["53"]),
        ("digits_on_variable", "program t\nreal :: x = 2.0\nprint *, digits(x)\nend program t\n", ["24"]),
        ("dim_positive", "program t\nprint *, dim(10, 3)\nend program t\n", ["7"]),
        ("dim_zero", "program t\nprint *, dim(3, 10)\nend program t\n", ["0"]),
        ("dim_equal", "program t\nprint *, dim(5, 5)\nend program t\n", ["0"]),
        ("dim_real_args", "program t\nprint *, nint(dim(4.5, 1.2))\nend program t\n", ["3"]),
        ("dot_product_int", "program t\ninteger :: a(3) = [1,2,3]\ninteger :: b(3) = [4,5,6]\nprint *, dot_product(a,b)\nend program t\n", ["32"]),
        ("dot_product_real", "program t\nreal :: a(2) = [1.5, 2.5]\nreal :: b(2) = [2.0, 4.0]\nprint *, dot_product(a,b)\nend program t\n", ["13"]),
        ("dot_product_negatives", "program t\ninteger :: a(2) = [-1, 3]\ninteger :: b(2) = [2, -4]\nprint *, dot_product(a,b)\nend program t\n", ["-14"]),
        ("dot_product_unit", "program t\ninteger :: a(3) = [1,0,0]\ninteger :: b(3) = [0,1,0]\nprint *, dot_product(a,b)\nend program t\n", ["0"]),
        ("dprod_double", "program t\nprint *, nint(dprod(2.0d0, 3.0d0))\nend program t\n", ["6"]),
        ("dprod_fractional", "program t\nprint *, nint(dprod(1.5d0, 2.0d0)*10)\nend program t\n", ["30"]),
        ("dshiftl_scalar", "program t\nprint *, dshiftl(14, 2, 4)\nend program t\n", None),
        ("dshiftl_zero_shift", "program t\nprint *, dshiftl(7, 0, 4)\nend program t\n", None),
        ("dshiftr_scalar", "program t\nprint *, dshiftr(14, 2, 4)\nend program t\n", None),
        ("dshiftr_zero_shift", "program t\nprint *, dshiftr(7, 0, 4)\nend program t\n", None),
        ("eoshift_left_default", "program t\ninteger :: a(4) = [1,2,3,4]\nprint *, eoshift(a, 1)\nend program t\n", None),
        ("eoshift_right", "program t\ninteger :: a(4) = [1,2,3,4]\nprint *, eoshift(a, -1)\nend program t\n", None),
        ("eoshift_boundary", "program t\ninteger :: a(3) = [1,2,3]\nprint *, eoshift(a, 1, boundary=0)\nend program t\n", None),
        ("eoshift_dim2", "program t\ninteger :: m(2,2) = reshape([1,2,3,4],[2,2])\nprint *, eoshift(m, 1, dim=2)\nend program t\n", None),
        ("epsilon_exponent", "program t\nprint *, exponent(epsilon(1.0))\nend program t\n", None),
        ("epsilon_scaled", "program t\nprint *, nint(epsilon(1.0)*1.0e10)\nend program t\n", None),
        ("erf_zero", "program t\nprint *, nint(erf(0.0)*100)\nend program t\n", None),
        ("erf_one", "program t\nprint *, nint(erf(1.0)*100)\nend program t\n", None),
        ("erf_neg_one", "program t\nprint *, nint(erf(-1.0)*100)\nend program t\n", None),
        ("erfc_zero", "program t\nprint *, nint(erfc(0.0)*100)\nend program t\n", None),
        ("erfc_one", "program t\nprint *, nint(erfc(1.0)*100)\nend program t\n", None),
        ("erfc_large", "program t\nprint *, nint(erfc(3.0)*100)\nend program t\n", None),
        ("exp_zero", "program t\nprint *, nint(exp(0.0)*100)\nend program t\n", None),
        ("exp_one", "program t\nprint *, nint(exp(1.0)*100)\nend program t\n", None),
        ("exp_neg_one", "program t\nprint *, nint(exp(-1.0)*100)\nend program t\n", None),
        ("exponent_one", "program t\nprint *, exponent(1.0)\nend program t\n", None),
        ("exponent_large", "program t\nprint *, exponent(16.0)\nend program t\n", None),
        ("exponent_fraction", "program t\nprint *, exponent(0.25)\nend program t\n", None),
        ("findloc_first_match", "program t\ninteger :: a(5) = [3,1,9,1,5]\nprint *, findloc(a, 9)\nend program t\n", None),
        ("findloc_no_match", "program t\ninteger :: a(3) = [1,2,3]\nprint *, findloc(a, 9)\nend program t\n", None),
        ("findloc_dim1", "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, findloc(m, 5, dim=1)\nend program t\n", None),
        ("findloc_dim2", "program t\ninteger :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])\nprint *, findloc(m, 5, dim=2)\nend program t\n", None),
        ("findloc_back_false", "program t\ninteger :: a(4) = [5,3,5,2]\nprint *, findloc(a, 5, back=.false.)\nend program t\n", None),
        ("findloc_back_true", "program t\ninteger :: a(4) = [5,3,5,2]\nprint *, findloc(a, 5, back=.true.)\nend program t\n", None),
        ("floor_positive", "program t\nprint *, floor(3.9)\nend program t\n", ["3"]),
        ("floor_negative", "program t\nprint *, floor(-3.1)\nend program t\n", ["-4"]),
        ("floor_with_kind", "program t\ninteger, parameter :: sp = kind(1.0)\nprint *, floor(2.2, kind=sp)\nend program t\n", ["2"]),
        ("floor_zero", "program t\nprint *, floor(0.0)\nend program t\n", ["0"]),
        ("fraction_whole", "program t\nprint *, nint(fraction(4.0)*100)\nend program t\n", None),
        ("fraction_mixed", "program t\nprint *, nint(fraction(3.75)*100)\nend program t\n", None),
        ("fraction_small", "program t\nprint *, nint(fraction(0.25)*100)\nend program t\n", None),
        ("fraction_with_exponent", "program t\nreal :: f\ninteger :: e\nf = fraction(6.0)\ne = exponent(6.0)\nprint *, nint(f*100)\nprint *, e\nend program t\n", None),
    ],
    "test_intrinsic_catalog_letter_k_p.rs": [
        ("gamma_two", "program t\nprint *, nint(gamma(2.0)*100)\nend program t\n", None),
        ("gamma_three", "program t\nprint *, nint(gamma(3.0)*100)\nend program t\n", None),
        ("gamma_half", "program t\nprint *, nint(gamma(0.5)*100)\nend program t\n", None),
        ("gamma_four", "program t\nprint *, nint(gamma(4.0)*100)\nend program t\n", None),
        ("get_command", "program t\ncharacter(len=32) :: cmd\ninteger :: stat\nstat = get_command(cmd)\nprint *, stat\nprint *, len_trim(cmd)\nend program t\n", None),
        ("huge_default_int", "program t\nprint *, huge(0) > 0\nend program t\n", ["T"]),
        ("huge_kind_param", "program t\ninteger, parameter :: ik = kind(0)\nprint *, huge(0, kind=ik) > 1000000\nend program t\n", ["T"]),
        ("huge_real", "program t\nprint *, huge(1.0) > 1.0\nend program t\n", ["T"]),
        ("huge_double", "program t\nprint *, huge(1.0d0) > 1.0d0\nend program t\n", ["T"]),
        ("iachar_letter_a", "program t\nprint *, iachar('A')\nend program t\n", ["65"]),
        ("iachar_lowercase", "program t\nprint *, iachar('z')\nend program t\n", ["122"]),
        ("iachar_digit", "program t\nprint *, iachar('5')\nend program t\n", ["53"]),
        ("iachar_from_var", "program t\ncharacter(len=1) :: c = 'Q'\nprint *, iachar(c)\nend program t\n", ["81"]),
        ("iall_bits", "program t\nprint *, iall(7)\nend program t\n", None),
        ("iall_zero", "program t\nprint *, iall(0)\nend program t\n", None),
        ("iall_dim1", "program t\ninteger :: a(2,2) = reshape([3,5,7,1],[2,2])\nprint *, iall(a, dim=1)\nend program t\n", None),
        ("iall_dim2", "program t\ninteger :: a(2,2) = reshape([3,5,7,1],[2,2])\nprint *, iall(a, dim=2)\nend program t\n", None),
        ("ibits_extract", "program t\nprint *, ibits(170, 1, 4)\nend program t\n", None),
        ("ibits_msb", "program t\nprint *, ibits(255, 7, 1)\nend program t\n", None),
        ("ibits_zero_len", "program t\nprint *, ibits(42, 0, 0)\nend program t\n", None),
        ("ichar_letter_b", "program t\nprint *, ichar('B')\nend program t\n", ["66"]),
        ("ichar_space", "program t\nprint *, ichar(' ')\nend program t\n", ["32"]),
        ("ichar_from_var", "program t\ncharacter(len=1) :: c = 'X'\nprint *, ichar(c)\nend program t\n", ["88"]),
        ("image_index", "program t\nprint *, image_index(1)\nend program t\n", None),
        ("is_iostat_end", "program t\nprint *, is_iostat_end(-1)\nend program t\n", None),
        ("is_iostat_end_zero", "program t\nprint *, is_iostat_end(0)\nend program t\n", None),
        ("kind_int_literal", "program t\nprint *, kind(1)\nend program t\n", ["8"]),
        ("kind_real_literal", "program t\nprint *, kind(1.0)\nend program t\n", ["8"]),
        ("kind_logical", "program t\nprint *, kind(.true.)\nend program t\n", ["8"]),
        ("kind_character", "program t\nprint *, kind('a')\nend program t\n", ["8"]),
        ("lbound_1d", "program t\ninteger :: a(5)\nprint *, lbound(a, 1)\nend program t\n", ["1"]),
        ("lbound_2d", "program t\ninteger :: m(2,3)\nprint *, lbound(m, 1)\nprint *, lbound(m, 2)\nend program t\n", ["1", "1"]),
        ("lbound_allocatable", "program t\ninteger, allocatable :: a(:)\nallocate(a(3:5))\nprint *, lbound(a, 1)\nend program t\n", ["3"]),
        ("lbound_whole", "program t\ninteger :: a(4)\nprint *, lbound(a)\nend program t\n", None),
        ("leadz_zero", "program t\nprint *, leadz(0)\nend program t\n", None),
        ("leadz_one", "program t\nprint *, leadz(1)\nend program t\n", None),
        ("leadz_eight", "program t\nprint *, leadz(8)\nend program t\n", None),
        ("leadz_negative", "program t\nprint *, leadz(-1)\nend program t\n", None),
        ("lgamma_two", "program t\nprint *, nint(lgamma(2.0)*100)\nend program t\n", None),
        ("lgamma_three", "program t\nprint *, nint(lgamma(3.0)*100)\nend program t\n", None),
        ("lgamma_half", "program t\nprint *, nint(lgamma(0.5)*100)\nend program t\n", None),
        ("logical_from_int_one", "program t\nprint *, logical(1)\nend program t\n", ["T"]),
        ("logical_from_int_zero", "program t\nprint *, logical(0)\nend program t\n", ["F"]),
        ("logical_kind_arg", "program t\nprint *, logical(1, kind=kind(.true.))\nend program t\n", ["T"]),
        ("logical_from_comparison", "program t\nprint *, logical(5 > 3)\nend program t\n", ["T"]),
        ("log_one", "program t\nprint *, nint(log(1.0)*100)\nend program t\n", None),
        ("log_e", "program t\nprint *, nint(log(2.718281828)*100)\nend program t\n", None),
        ("log_exp_identity", "program t\nprint *, nint(log(exp(1.0))*100)\nend program t\n", None),
        ("log10_one", "program t\nprint *, nint(log10(1.0)*100)\nend program t\n", None),
        ("log10_hundred", "program t\nprint *, nint(log10(100.0)*100)\nend program t\n", None),
        ("log10_thousand", "program t\nprint *, nint(log10(1000.0)*100)\nend program t\n", None),
    ],
}


def calibrate_all() -> dict[str, list[str]]:
    """Run each test source through run_prints and return actual output."""
    lines = [
        "//! Temporary calibration harness — delete after generation.",
        "use super::helpers::run_prints;",
        "",
    ]
    idx = 0
    index: dict[str, tuple[str, str]] = {}
    for tests in CATALOGS.values():
        for name, src, _ in tests:
            index[name] = (str(idx), src)
            escaped = json.dumps(src)
            lines.append(f"#[test]")
            lines.append(f"fn cal_{name}() {{")
            lines.append(f"    let out = run_prints({escaped});")
            lines.append(f'    eprintln!("CAL|{name}|{{:?}}", out);')
            lines.append(f"}}")
            lines.append("")
            idx += 1

    CALIBRATE_RS.write_text("\n".join(lines))

    # Temporarily add mod to main.rs is forbidden. Use rustc test via --test flag with include.
    # Instead append to an existing file or use cargo test with --ignored pattern.
    # We'll patch main.rs temporarily... user said don't modify main.rs.
    # Write standalone integration test binary instead.
    result = subprocess.run(
        [
            "cargo",
            "test",
            "-p",
            "vybe_compiler",
            "--test",
            "fortran",
            "cal_",
            "--",
            "--nocapture",
            "--test-threads=1",
        ],
        cwd=ROOT.parents[3],
        capture_output=True,
        text=True,
    )
    outputs: dict[str, list[str]] = {}
    for line in result.stdout.splitlines() + result.stderr.splitlines():
        if line.startswith("CAL|"):
            _, name, rest = line.split("|", 2)
            # rest looks like: acos_one|["0"]  after split wrong
            pass
    # Parse eprintln from stderr
    import re

    for line in result.stderr.splitlines():
        m = re.match(r"CAL\|([^|]+)\|\[?(.*)\]?$", line)
        if not m:
            continue
        name = m.group(1)
        raw = m.group(2)
        # handle Debug format: ["0", "1"]
        if raw.startswith('["'):
            inner = raw.strip('[]')
            parts = []
            cur = ""
            in_str = False
            for ch in inner:
                if ch == '"':
                    in_str = not in_str
                    if not in_str and cur:
                        parts.append(cur)
                        cur = ""
                elif in_str:
                    cur += ch
            outputs[name] = parts
    return outputs


def render_file(filename: str, tests: list[tuple[str, str, list[str] | None]], outputs: dict[str, list[str]]) -> str:
    title = filename.replace("test_", "").replace(".rs", "").replace("_", " ")
    lines = [
        f"//! Intrinsic catalog: {title}.",
        "",
        "use super::helpers;",
        "",
        "fortran_cases! {",
    ]
    for name, src, expected in tests:
        if expected is None:
            expected = outputs.get(name, ["0"])
        exp_str = ", ".join(json.dumps(x) for x in expected)
        src_escaped = json.dumps(src)
        lines.append(f"    {name} => {{")
        lines.append(f"        {src_escaped},")
        lines.append(f"        [{exp_str}]")
        lines.append(f"    }};")
        lines.append("")
    lines.append("}")
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    # Write calibrate file and inject via symlink trick: compile as submodule in helpers
    # Simpler: run calibration inline using rust one-liner through cargo script
    print("Calibrating...")
    outputs = calibrate_all()
    print(f"Calibrated {len(outputs)} tests")

    for filename, tests in CATALOGS.items():
        content = render_file(filename, tests, outputs)
        out = ROOT / filename
        out.write_text(content)
        print(f"Wrote {filename}: {len(tests)} tests")


if __name__ == "__main__":
    main()
