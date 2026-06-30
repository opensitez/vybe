//! Temporary probe — delete after generation.
use super::helpers::run_prints;

macro_rules! probe {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            eprintln!("{}: {:?}", stringify!($name), out);
        }
    };
}

probe!(p_maskl4, "program t\nprint *, popcount(maskl(4))\nend program t\n");
probe!(p_maskr4, "program t\nprint *, popcount(maskr(4))\nend program t\n");
probe!(p_maxexp, "program t\nprint *, maxexponent(1.0)\nend program t\n");
probe!(p_minexp, "program t\nprint *, minexponent(1.0)\nend program t\n");
probe!(p_norm2, "program t\nprint *, nint(norm2([3.0,4.0,0.0]))\nend program t\n");
probe!(p_nint, "program t\nprint *, nint(3.7)\nend program t\n");
probe!(p_sind90, "program t\nprint *, nint(sind(90.0))\nend program t\n");
probe!(p_cosd0, "program t\nprint *, nint(cosd(0.0))\nend program t\n");
probe!(p_tand45, "program t\nprint *, nint(tand(45.0))\nend program t\n");
probe!(p_asind, "program t\nprint *, nint(asind(0.5))\nend program t\n");
probe!(p_shifta, "program t\nprint *, shifta(8, 1)\nend program t\n");
probe!(p_shiftl, "program t\nprint *, shiftl(1, 3)\nend program t\n");
probe!(p_shiftr, "program t\nprint *, shiftr(8, 1)\nend program t\n");
probe!(p_popcnt, "program t\nprint *, popcount(255)\nend program t\n");
probe!(p_poppar, "program t\nprint *, poppar(255)\nend program t\n");
probe!(p_parity, "program t\nlogical :: a(2)=[.true.,.true.]\nprint *, merge(1,0,parity(a))\nend program t\n");
probe!(p_shape, "program t\ninteger :: a(2,3)\nprint *, shape(a,1)\nprint *, shape(a,2)\nend program t\n");
probe!(p_ubound, "program t\ninteger :: a(3,4)\nprint *, ubound(a,1)\nprint *, ubound(a,2)\nend program t\n");
probe!(p_repeat, "program t\nprint *, len(repeat('ab',3))\nend program t\n");
probe!(p_trim, "program t\nprint *, len(trim(' hi '))\nend program t\n");
probe!(p_verify, "program t\nprint *, verify('abc','b')\nend program t\n");
probe!(p_scan, "program t\nprint *, scan('abc','c')\nend program t\n");
probe!(p_spread, "program t\ninteger :: a(3)=[1,2,3]\ninteger :: b(2,3)\nb=spread(a,dim=1,n=2)\nprint *, b(1,2)\nprint *, b(2,2)\nend program t\n");
probe!(p_sign, "program t\nprint *, sign(5,-1)\nend program t\n");
probe!(p_sqrt, "program t\nprint *, nint(sqrt(16.0))\nend program t\n");
probe!(p_sinh0, "program t\nprint *, nint(sinh(0.0)*100)\nend program t\n");
probe!(p_sin0, "program t\nprint *, nint(sin(0.0)*100)\nend program t\n");
probe!(p_scale, "program t\nprint *, nint(scale(1.0,2))\nend program t\n");
probe!(p_setexp, "program t\nprint *, nint(set_exponent(1.0,3))\nend program t\n");
probe!(p_radix, "program t\nprint *, radix(1.0)\nend program t\n");
probe!(p_range, "program t\nprint *, range(0)\nend program t\n");
probe!(p_precision, "program t\nprint *, precision(1.0)\nend program t\n");
probe!(p_storage, "program t\nprint *, storage_size(0)\nend program t\n");
probe!(p_selected_int, "program t\nprint *, selected_int_kind(9)\nend program t\n");
probe!(p_selected_real, "program t\nprint *, selected_real_kind(6)\nend program t\n");
probe!(p_pack, "program t\ninteger :: a(4)=[1,2,3,4]\nlogical :: m(4)=[.true.,.false.,.true.,.false.]\ninteger :: b(2)\nb=pack(a,m)\nprint *, b(1)\nprint *, b(2)\nend program t\n");
probe!(p_unpack, "program t\ninteger :: a(2)=[7,9]\nlogical :: m(4)=[.true.,.false.,.true.,.false.]\ninteger :: b(4)\nb=unpack(a,m,0)\nprint *, b(1)\nprint *, b(3)\nend program t\n");
probe!(p_reshape, "program t\ninteger :: a(6)=[1,2,3,4,5,6]\ninteger :: b(2,3)\nb=reshape(a,[2,3],order=[2,1])\nprint *, b(1,1)\nprint *, b(2,1)\nend program t\n");
probe!(p_product, "program t\ninteger :: m(2,3)=reshape([1,2,3,4,5,6],[2,3])\nprint *, product(m,dim=1)\nprint *, product(m,dim=1)\nend program t\n");
probe!(p_sumdim, "program t\ninteger :: m(2,3)=reshape([1,2,3,4,5,6],[2,3])\nprint *, sum(m,dim=2)\nprint *, sum(m,dim=2)\nend program t\n");
probe!(p_present, "program t\ncall sub(1)\ncontains\nsubroutine sub(x,optional y)\ninteger,intent(in)::x\ninteger,optional,intent(in)::y\nprint *, merge(1,0,.not.present(y))\nend subroutine sub\nend program t\n");
probe!(p_dim, "program t\nprint *, nint(dim(10.5,3.2))\nend program t\n");
probe!(p_atan2, "program t\nprint *, nint(atan2(2.0,2.0)*1000)\nend program t\n");
probe!(p_matmul4, "program t\ninteger :: a(4,4),b(4,4),c(4,4)\na=reshape([(i,i=1,16)],[4,4])\nb=0;b(1,1)=1;b(2,2)=1;b(3,3)=1;b(4,4)=1\nc=matmul(a,b)\nprint *, c(1,1)\nprint *, c(4,4)\nend program t\n");
probe!(p_nearest, "program t\nprint *, merge(1,0,nearest(1.0,1.0)>1.0)\nend program t\n");
probe!(p_newline, "program t\ncharacter(len=1)::nl\nnl=new_line('a')\nprint *, ichar(nl)\nend program t\n");
probe!(p_numimg, "program t\nprint *, num_images()\nend program t\n");
probe!(p_null, "program t\nuse iso_c_binding\nprint *, merge(1,0,c_associated(c_null_ptr))\nend program t\n");
probe!(p_sysclock, "program t\ninteger :: c,r,m\ncall system_clock(c,r,m)\nprint *, merge(1,0,m>0)\nend program t\n");
probe!(p_trailz8, "program t\nprint *, trailz(8)\nend program t\n");
probe!(p_tan0, "program t\nprint *, nint(tan(0.0)*100)\nend program t\n");
probe!(p_tanh0, "program t\nprint *, nint(tanh(0.0)*100)\nend program t\n");
probe!(p_tiny, "program t\nprint *, merge(1,0,tiny(1.0)>0.0)\nend program t\n");
probe!(p_acosd, "program t\nprint *, nint(acosd(0.5))\nend program t\n");
probe!(p_atand, "program t\nprint *, nint(atand(1.0))\nend program t\n");
probe!(p_merge3, "program t\nprint *, merge(10,20,.true.)\nend program t\n");
probe!(p_merge_arr, "program t\ninteger :: a(3)=[1,2,3]\ninteger :: b(3)=[9,8,7]\nlogical :: m(3)=[.true.,.false.,.true.]\nprint *, merge(a,b,m)\nprint *, merge(a,b,m)\nend program t\n");
