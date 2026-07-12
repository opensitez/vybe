use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(assign_move_01, "program t\ninteger :: a=1,b\nb=a\nend program t\n");
c!(assign_move_lit_02, "program t\ninteger :: b\nb=1\nend program t\n");
c!(assign_char_03, "program t\ncharacter(len=3) :: a='abc', b\nb=a\nend program t\n");
c!(assign_array_04, "program t\ninteger :: a(3)=[1,2,3], b(3)\nb=a\nend program t\n");
c!(assign_slice_05, "program t\ninteger :: a(4)=[1,2,3,4], b(2)\nb=a(2:3)\nend program t\n");
c!(assign_pointer_06, "program t\ninteger,target :: x\ninteger,pointer :: p\np => x\np = 1\nend program t\n");
c!(assign_compute_07, "program t\ninteger :: x\nx = 1 + 2\nend program t\n");
c!(assign_compute_paren_08, "program t\ninteger :: x\nx = (1 + 2) * 3\nend program t\n");
c!(assign_compute_real_09, "program t\nreal :: x\nx = 1.5 + 2.5\nend program t\n");
c!(assign_logical_10, "program t\nlogical :: x\nx = .true.\nend program t\n");
c!(assign_complex_11, "program t\ncomplex :: z\nz = (1.0, 2.0)\nend program t\n");
c!(assign_kind_conv_12, "program t\ninteger :: i\nreal :: r=1.5\ni = int(r)\nend program t\n");
c!(assign_char_concat_13, "program t\ncharacter(len=6) :: s\ns = 'ab'//'cd'\nend program t\n");
c!(assign_substring_14, "program t\ncharacter(len=5) :: s='hello'\ns(1:2)='HE'\nend program t\n");
c!(assign_structure_15, "type :: t\n integer :: x\nend type t\nprogram p\ntype(t) :: a,b\na%x=1\nb=a\nend program p\n");
c!(assign_constructor_16, "type :: t\n integer :: x\nend type t\nprogram p\ntype(t) :: a\na = t(1)\nend program p\n");
c!(assign_masked_where_17, "program t\ninteger :: a(3)=[1,2,3]\nwhere (a > 1) a = 0\nend program t\n");
c!(assign_forall_18, "program t\ninteger :: a(3)\nforall(i=1:3) a(i)=i\nend program t\n");
c!(assign_component_19, "type :: t\n integer :: x\nend type t\nprogram p\ntype(t) :: a\na%x = 2\nend program p\n");
c!(assign_alloc_comp_20, "type :: t\n integer, allocatable :: a(:)\nend type t\nprogram p\ntype(t) :: x\nallocate(x%a(2))\nx%a = [1,2]\nend program p\n");
c!(assign_pointer_array_21, "program t\ninteger,target :: a(2)=[1,2]\ninteger,pointer :: p(:)\np => a\np = [3,4]\nend program t\n");
c!(assign_do_index_22, "program t\ninteger :: i, a(3)\ndo i=1,3\n a(i)=i\nend do\nend program t\n");
c!(assign_parameter_expr_23, "program t\ninteger, parameter :: a=1+2\nprint *, a\nend program t\n");
c!(assign_spec_expr_24, "program t\ninteger, parameter :: n=3\ninteger :: a(n)\nprint *, size(a)\nend program t\n");
c!(assign_reshape_25, "program t\ninteger :: a(2,2)\na = reshape([1,2,3,4],[2,2])\nend program t\n");
c!(assign_transfer_26, "program t\ninteger :: x\nx = transfer('abcd', 1)\nend program t\n");
c!(assign_merge_27, "program t\ninteger :: x\nx = merge(1,2,.true.)\nend program t\n");
c!(assign_spread_28, "program t\ninteger :: a(2,2)\na = spread([1,2], dim=2, ncopies=2)\nend program t\n");
c!(assign_pack_29, "program t\ninteger :: a(3)=[1,2,3], b(2)\nb = pack(a, a>1)\nend program t\n");
c!(assign_unpack_30, "program t\ninteger :: a(3)=[1,2,3], b(3), v(2)=[9,8]\nb = unpack(v, a>1, a)\nend program t\n")