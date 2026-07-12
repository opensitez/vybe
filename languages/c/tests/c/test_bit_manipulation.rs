use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    set_bit => {
        body: "int x = 0;\nx |= (1 << 3);\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["8"]
    },
    clear_bit => {
        body: "int x = 0xFF;\nx &= ~(1 << 4);\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["239"]
    },
    toggle_bit => {
        body: "int x = 0b1010;\nx ^= 0b0110;\nprintf(\"%d\\n\", x);\nreturn 0;",
        expect: ["12"]
    },
    test_bit => {
        body: "int x = 0b10110;\nint bit3 = (x >> 2) & 1;\nprintf(\"%d\\n\", bit3);\nreturn 0;",
        expect: ["1"]
    },
    count_set_bits => {
        body: r#"
int n = 0b10110110;
int count = 0;
while (n) { count += n & 1; n >>= 1; }
printf("%d\n", count);
return 0;
"#,
        expect: ["5"]
    },
    is_power_of_two => {
        body: "int x = 64;\nprintf(\"%d\\n\", (x & (x - 1)) == 0 ? 1 : 0);\nreturn 0;",
        expect: ["1"]
    },
    is_not_power_of_two => {
        body: "int x = 60;\nprintf(\"%d\\n\", (x & (x - 1)) == 0 ? 1 : 0);\nreturn 0;",
        expect: ["0"]
    },
    extract_nibble => {
        body: "unsigned int x = 0xABCD;\nunsigned int nibble = (x >> 8) & 0xF;\nprintf(\"%x\\n\", nibble);\nreturn 0;",
        expect: ["b"]
    },
    swap_bytes => {
        body: "unsigned short x = 0x1234;\nunsigned short swapped = ((x & 0xFF) << 8) | ((x >> 8) & 0xFF);\nprintf(\"%x\\n\", swapped);\nreturn 0;",
        expect: ["3412"]
    },
    bit_rotation_left => {
        body: "unsigned int x = 0x12345678;\nunsigned int rot = (x << 4) | (x >> 28);\nprintf(\"%x\\n\", rot);\nreturn 0;",
        expect: ["23456781"]
    }
}
