use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(and_basic, "System.out.println(6 & 3);", "2");
jt!(or_basic, "System.out.println(6 | 3);", "7");
jt!(xor_basic, "System.out.println(6 ^ 3);", "5");
jt!(not_basic, "System.out.println(~0);", "-1");
jt!(not_non_zero, "System.out.println(~1);", "-2");
jt!(not_negative, "System.out.println(~(-2));", "1");
jt!(left_shift_one, "System.out.println(1 << 3);", "8");
jt!(left_shift_two, "System.out.println(3 << 1);", "6");
jt!(right_shift_positive, "System.out.println(8 >> 2);", "2");
jt!(right_shift_negative, "System.out.println((-8) >> 2);", "-2");
jt!(
    unsigned_right_shift_positive,
    "System.out.println(8 >>> 2);",
    "2"
);
jt!(
    unsigned_right_shift_negative,
    "System.out.println((-8) >>> 30);",
    "3"
);
jt!(bit_clear, "int v = 7; v &= ~2; System.out.println(v);", "5");
jt!(bit_set, "int v = 1; v |= 8; System.out.println(v);", "9");
jt!(bit_toggle, "int v = 5; v ^= 1; System.out.println(v);", "4");
jt!(
    mixed_bitwise_ops,
    "int v = 7; int w = (v << 1) & 10; System.out.println(w);",
    "6"
);
jt!(
    bit_mask_and_shift,
    "int v = 1; v = (v << 4) | 3; v &= 19; System.out.println(v);",
    "19"
);
jt!(
    multiple_shifts,
    "int v = 1; int r = (v << 8) >> 2; System.out.println(r);",
    "64"
);
jt!(
    sign_bit_check,
    "int v = Integer.MIN_VALUE; System.out.println((v & Integer.MIN_VALUE) != 0);",
    "true"
);
jt!(low_bit_check, "System.out.println((3 & 1) != 0);", "true");
jt!(
    low_bit_check_false,
    "System.out.println((2 & 1) != 0);",
    "false"
);
jt!(
    compose_parity,
    "int v = 10; System.out.println((v & 1) == 0);",
    "true"
);
jt!(
    xor_swap_like,
    r#"int a = 1; int b = 3; a ^= b; b ^= a; a ^= b; System.out.println(a + "," + b);"#,
    "3,1"
);
jt!(
    bitwise_or_chain,
    "System.out.println((1 | 2 | 4 | 8));",
    "15"
);
jt!(bitwise_and_chain, "System.out.println((15 & 14 & 7));", "6");
jt!(bitwise_xor_chain, "System.out.println((15 ^ 10 ^ 3));", "4");
jt!(shift_then_mask, "System.out.println((1 << 3) & 15);", "8");
jt!(
    mask_parity,
    "int v = 13; System.out.println((v & 2) == 0);",
    "false"
);
jt!(
    ternary_with_bitops,
    "int v = 2; System.out.println((v & 1) == 0 ? 0 : 1);",
    "0"
);
jt!(
    bitwise_addition_flags,
    "int a = 1; a <<= 2; a |= 1; System.out.println(a);",
    "5"
);
jt!(
    double_shift_back,
    "int a = 12; a >>= 2; a <<= 2; System.out.println(a);",
    "8"
);
