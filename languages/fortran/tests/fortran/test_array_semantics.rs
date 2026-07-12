use super::helpers::compile_ok;
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    arr_assumed_shape_01,
    "subroutine s(a)
integer::a(:)
end subroutine s
"
);
c!(
    arr_assumed_size_02,
    "subroutine s(a)
integer::a(*)
end subroutine s
"
);
c!(
    arr_explicit_shape_03,
    "subroutine s(a)
integer::a(3)
end subroutine s
"
);
c!(
    arr_deferred_shape_04,
    "program p
integer, allocatable::a(:)
allocate(a(3))
end program p
"
);
c!(
    arr_constructor_05,
    "program p
integer::a(3)
a=[1,2,3]
print *,a
end program p
"
);
c!(
    arr_assign_06,
    "program p
integer::a(3),b(3)
a=[1,2,3]
b=a
print *,b
end program p
"
);
c!(
    arr_temp_07,
    "program p
integer::a(3)=[1,2,3]
print *,a+1
end program p
"
);
c!(
    arr_conform_08,
    "program p
integer::a(3)=[1,2,3],b(3)=[4,5,6]
print *,a+b
end program p
"
);
c!(
    arr_vector_sub_09,
    "program p
integer::a(4)=[1,2,3,4],i(2)=[1,3]
print *,a(i)
end program p
"
);
c!(
    arr_stride_10,
    "program p
integer::a(5)=[1,2,3,4,5]
print *,a(1:5:2)
end program p
"
);
c!(
    arr_zero_size_11,
    "program p
integer, allocatable::a(:)
allocate(a(0))
print *,size(a)
end program p
"
);
c!(
    arr_bounds_12,
    "program p
integer::a(0:2)
print *,lbound(a),ubound(a)
end program p
"
);
c!(
    arr_lower_13,
    "program p
integer::a(-1:1)
print *,lbound(a)
end program p
"
);
c!(
    arr_upper_14,
    "program p
integer::a(2:4)
print *,ubound(a)
end program p
"
);
c!(
    arr_reshape_15,
    "program p
integer::a(2,2)
a=reshape([1,2,3,4],[2,2])
print *,a
end program p
"
);
c!(
    arr_section_assign_16,
    "program p
integer::a(4)=[1,2,3,4]
a(2:3)=0
print *,a
end program p
"
);
c!(
    arr_mask_where_17,
    "program p
integer::a(3)=[1,2,3]
where(a>1) a=a+1
print *,a
end program p
"
);
c!(
    arr_pack_18,
    "program p
integer::a(3)=[1,2,3]
print *,pack(a,a>1)
end program p
"
);
c!(
    arr_unpack_19,
    "program p
integer::a(2)=[1,2],f(3)=[0,0,0]
print *,unpack(a,[.true.,.false.,.true.],f)
end program p
"
);
c!(
    arr_matmul_20,
    "program p
integer::a(2,2)=reshape([1,2,3,4],[2,2]),b(2,2)=reshape([1,0,0,1],[2,2])
print *,matmul(a,b)
end program p
"
);
