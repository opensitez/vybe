use super::helpers::compile_ok;

macro_rules! c { ($name:ident, $src:expr) => { #[test] fn $name() { compile_ok($src); } }; }

c!(sort_call_01, "program t\ninteger :: a(3)=[3,1,2]\ncall sort(a)\nend program t\n");
c!(merge_call_02, "program t\ninteger :: a(2)=[1,3], b(2)=[2,4]\nprint *, merge(a(1), b(1), .true.)\nend program t\n");
c!(pack_03, "program t\ninteger :: a(3)=[1,2,3], b(2)\nb = pack(a, a>1)\nend program t\n");
c!(unpack_04, "program t\ninteger :: a(3)=[1,2,3], b(3), v(2)=[9,8]\nb = unpack(v, a>1, a)\nend program t\n");
c!(reshape_05, "program t\ninteger :: a(2,2)\na = reshape([1,2,3,4],[2,2])\nend program t\n");
c!(cshift_06, "program t\ninteger :: a(3)=[1,2,3]\na = cshift(a,1)\nend program t\n");
c!(eoshift_07, "program t\ninteger :: a(3)=[1,2,3]\na = eoshift(a,1)\nend program t\n");
c!(spread_08, "program t\ninteger :: a(2,2)\na = spread([1,2], dim=2, ncopies=2)\nend program t\n");
c!(transpose_09, "program t\ninteger :: a(2,2), b(2,2)\na = reshape([1,2,3,4],[2,2])\nb = transpose(a)\nend program t\n");
c!(matmul_10, "program t\ninteger :: a(2,2), b(2,2), c(2,2)\na = 1\nb = 2\nc = matmul(a,b)\nend program t\n");
c!(dot_product_11, "program t\ninteger :: a(3)=[1,2,3], b(3)=[4,5,6], c\nc = dot_product(a,b)\nend program t\n");
c!(sum_12, "program t\ninteger :: a(3)=[1,2,3]\nprint *, sum(a)\nend program t\n");
c!(product_13, "program t\ninteger :: a(3)=[1,2,3]\nprint *, product(a)\nend program t\n");
c!(maxval_14, "program t\ninteger :: a(3)=[1,2,3]\nprint *, maxval(a)\nend program t\n");
c!(minval_15, "program t\ninteger :: a(3)=[1,2,3]\nprint *, minval(a)\nend program t\n");
c!(maxloc_16, "program t\ninteger :: a(3)=[1,3,2]\nprint *, maxloc(a)\nend program t\n");
c!(minloc_17, "program t\ninteger :: a(3)=[1,3,2]\nprint *, minloc(a)\nend program t\n");
c!(findloc_18, "program t\ninteger :: a(3)=[1,3,2]\nprint *, findloc(a, 3)\nend program t\n");
c!(count_19, "program t\nlogical :: m(3)=[.true.,.false.,.true.]\nprint *, count(m)\nend program t\n");
c!(any_20, "program t\nlogical :: m(3)=[.true.,.false.,.true.]\nprint *, any(m)\nend program t\n");
c!(all_21, "program t\nlogical :: m(3)=[.true.,.true.,.true.]\nprint *, all(m)\nend program t\n");
c!(parity_22, "program t\nlogical :: m(3)=[.true.,.false.,.true.]\nprint *, parity(m)\nend program t\n");
c!(merge_bits_23, "program t\ninteger :: x\nx = merge_bits(1,2,3)\nend program t\n");
c!(maskl_24, "program t\ninteger :: x\nx = maskl(3)\nend program t\n");
c!(maskr_25, "program t\ninteger :: x\nx = maskr(3)\nend program t\n");
c!(ishft_26, "program t\ninteger :: x\nx = ishft(1,2)\nend program t\n");
c!(ibset_27, "program t\ninteger :: x\nx = ibset(0,1)\nend program t\n");
c!(ibclr_28, "program t\ninteger :: x\nx = ibclr(3,0)\nend program t\n");
c!(btest_29, "program t\nlogical :: x\nx = btest(3,0)\nprint *, x\nend program t\n");
c!(merge_intrinsic_30, "program t\ninteger :: x\nx = merge(1,2,.false.)\nend program t\n")