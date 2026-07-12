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
    ieee_arith_01,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_datatype(1.0)
end program p
"
);
c!(
    ieee_features_02,
    "program p
use, intrinsic :: ieee_features
print *, ieee_support_underflow_control(1.0)
end program p
"
);
c!(
    ieee_except_03,
    "program p
use, intrinsic :: ieee_exceptions
logical :: l
call ieee_get_halting_mode(ieee_divide_by_zero, l)
end program p
"
);
c!(
    ieee_round_04,
    "program p
use, intrinsic :: ieee_arithmetic
logical :: l
call ieee_get_rounding_mode(l)
end program p
"
);
c!(
    ieee_flags_05,
    "program p
use, intrinsic :: ieee_exceptions
logical :: l
call ieee_get_flag(ieee_overflow, l)
end program p
"
);
c!(
    ieee_classify_06,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_class(1.0)
end program p
"
);
c!(
    ieee_copy_sign_07,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_copy_sign(1.0,-2.0)
end program p
"
);
c!(
    ieee_next_after_08,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_next_after(1.0,2.0)
end program p
"
);
c!(
    ieee_scalb_09,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_scalb(1.0,2)
end program p
"
);
c!(
    ieee_is_nan_10,
    "program p
use, intrinsic :: ieee_arithmetic
real :: x
x = 0.0/0.0
print *, ieee_is_nan(x)
end program p
"
);
c!(
    ieee_is_finite_11,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_is_finite(1.0)
end program p
"
);
c!(
    ieee_is_normal_12,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_is_normal(1.0)
end program p
"
);
c!(
    ieee_value_13,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_value(1.0, ieee_positive_inf)
end program p
"
);
c!(
    ieee_support_inf_14,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_inf(1.0)
end program p
"
);
c!(
    ieee_support_nan_15,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_nan(1.0)
end program p
"
);
c!(
    ieee_rem_16,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_rem(7.0,3.0)
end program p
"
);
c!(
    ieee_rint_17,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_rint(1.6)
end program p
"
);
c!(
    ieee_logb_18,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_logb(8.0)
end program p
"
);
c!(
    ieee_unordered_19,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_unordered(1.0, 2.0)
end program p
"
);
c!(
    ieee_datatype_20,
    "program p
use, intrinsic :: ieee_arithmetic
print *, ieee_support_datatype((1.0,2.0))
end program p
"
);
