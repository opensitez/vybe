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
    num_intk_01,
    "program p
integer(kind=4)::x
print *,1
end program p
"
);
c!(
    num_realk_02,
    "program p
real(kind=8)::x
print *,1
end program p
"
);
c!(
    num_complex_03,
    "program p
complex(kind=8)::z
print *,1
end program p
"
);
c!(
    num_conv_04,
    "program p
integer::i
real::r=1.5
i=int(r)
print *,i
end program p
"
);
c!(
    num_prec_05,
    "program p
print *, selected_real_kind(6)
end program p
"
);
c!(
    num_radix_06,
    "program p
print *, radix(1.0)
end program p
"
);
c!(
    num_round_07,
    "program p
print *, nint(1.6)
end program p
"
);
c!(
    num_nan_08,
    "program p
real::x
x=0.0/0.0
print *,x
end program p
"
);
c!(
    num_inf_09,
    "program p
real::x
x=1.0/0.0
print *,x
end program p
"
);
c!(
    num_signed_zero_10,
    "program p
real::x=-0.0
print *,x
end program p
"
);
c!(
    num_subnormal_11,
    "program p
print *, tiny(1.0)
end program p
"
);
c!(
    num_model_12,
    "program p
print *, digits(1.0)
end program p
"
);
c!(
    num_add_13,
    "program p
integer::a=1,b=2
print *,a+b
end program p
"
);
c!(
    num_sub_14,
    "program p
integer::a=5,b=2
print *,a-b
end program p
"
);
c!(
    num_mul_15,
    "program p
integer::a=3,b=4
print *,a*b
end program p
"
);
c!(
    num_div_16,
    "program p
real::a=8.0,b=2.0
print *,a/b
end program p
"
);
c!(
    num_mod_17,
    "program p
print *,mod(7,3)
end program p
"
);
c!(
    num_modulo_18,
    "program p
print *,modulo(-7,3)
end program p
"
);
c!(
    num_sign_19,
    "program p
print *,sign(2.0,-1.0)
end program p
"
);
c!(
    num_dim_20,
    "program p
print *,dim(5,3)
end program p
"
);
