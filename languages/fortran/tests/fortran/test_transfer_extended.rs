//! Extended TRANSFER intrinsic: integer/real bit reinterpretation, character
//! encodings, SIZE truncation/expansion, and array-level transfers with runtime
//! checks (roundtrips, ichar, byte extraction).

fortran_cases! {
    // ── Integer same-type scalar reinterpretation ─────────────────────

    transfer_scalar_zero_is_zero => {
        "program t\ninteger :: i = 0\nprint *, transfer(i, 0)\nend program t\n",
        ["0"]
    };

    transfer_scalar_minus_one => {
        "program t\ninteger :: i = -1\nprint *, transfer(i, 0)\nend program t\n",
        ["-1"]
    };

    transfer_scalar_fortytwo => {
        "program t\ninteger :: i = 42\nprint *, transfer(i, 0)\nend program t\n",
        ["42"]
    };

    transfer_scalar_million => {
        "program t\ninteger :: i = 1000000\nprint *, transfer(i, 0)\nend program t\n",
        ["1000000"]
    };

    transfer_scalar_hex_pattern => {
        "program t\ninteger :: i = 305419896\nprint *, transfer(i, 0)\nend program t\n",
        ["305419896"]
    };

    transfer_scalar_negative_nine_nine_nine => {
        "program t\ninteger :: i = -999\nprint *, transfer(i, 0)\nend program t\n",
        ["-999"]
    };

    transfer_scalar_kind1_byte => {
        "program t\ninteger(kind=1) :: b = 127_1\ninteger :: n\nn = transfer(b, 0)\nprint *, n\nend program t\n",
        ["127"]
    };

    transfer_scalar_kind2_short => {
        "program t\ninteger(kind=2) :: s = 32000_2\ninteger :: n\nn = transfer(s, 0)\nprint *, n\nend program t\n",
        ["32000"]
    };

    // ── Integer to real bit pattern roundtrips ──────────────────────────

    transfer_int_real_roundtrip_zero => {
        "program t\ninteger :: i = 0, j\nreal :: r\nr = transfer(i, 0.0)\nj = transfer(r, 0)\nprint *, j\nend program t\n",
        ["0"]
    };

    transfer_int_real_roundtrip_one => {
        "program t\ninteger :: i = 1, j\nreal :: r\nr = transfer(i, 0.0)\nj = transfer(r, 0)\nprint *, j\nend program t\n",
        ["1"]
    };

    transfer_int_real_roundtrip_fortytwo => {
        "program t\ninteger :: i = 42, j\nreal :: r\nr = transfer(i, 0.0)\nj = transfer(r, 0)\nprint *, j\nend program t\n",
        ["42"]
    };

    transfer_int_real_roundtrip_negative => {
        "program t\ninteger :: i = -7, j\nreal :: r\nr = transfer(i, 0.0)\nj = transfer(r, 0)\nprint *, j\nend program t\n",
        ["-7"]
    };

    transfer_int_real_roundtrip_large => {
        "program t\ninteger :: i = 65536, j\nreal :: r\nr = transfer(i, 0.0)\nj = transfer(r, 0)\nprint *, j\nend program t\n",
        ["65536"]
    };

    transfer_kind8_int_real_roundtrip => {
        "program t\ninteger(kind=8) :: i = 9876543210_8, j\nreal(kind=8) :: r\nr = transfer(i, 0.0d0)\nj = transfer(r, 0_8)\nprint *, j\nend program t\n",
        ["9876543210"]
    };

    transfer_int_to_real_then_back_equality => {
        "program t\ninteger :: i = 12345, j\nreal :: r\nr = transfer(i, 0.0)\nj = transfer(r, 0)\nprint *, i == j\nend program t\n",
        ["1"]
    };

    transfer_real_bits_from_integer_one => {
        "program t\ninteger :: i = 1\nreal :: r\nr = transfer(i, 0.0)\nprint *, transfer(r, 0)\nend program t\n",
        ["1"]
    };

    // ── Real to integer bit pattern roundtrips ──────────────────────────

    transfer_real_int_roundtrip_zero => {
        "program t\nreal :: x = 0.0\ninteger :: n\nn = transfer(x, 0)\nprint *, transfer(n, 0.0) == x\nend program t\n",
        ["1"]
    };

    transfer_real_int_roundtrip_one => {
        "program t\nreal :: x = 1.0\ninteger :: n\nn = transfer(x, 0)\nprint *, transfer(n, 0.0) == x\nend program t\n",
        ["1"]
    };

    transfer_real_int_roundtrip_pi => {
        "program t\nreal :: x = 3.14\ninteger :: n\nn = transfer(x, 0)\nprint *, transfer(n, 0.0) == x\nend program t\n",
        ["1"]
    };

    transfer_real_negative_bits_roundtrip => {
        "program t\nreal :: x = -2.5\ninteger :: n\nn = transfer(x, 0)\nprint *, transfer(n, 0.0) == x\nend program t\n",
        ["1"]
    };

    // ── Character to integer and back ─────────────────────────────────

    transfer_char_a_to_integer_le => {
        "program t\ncharacter(len=1) :: c = 'A'\nprint *, transfer(c, 0)\nend program t\n",
        ["65"]
    };

    transfer_char_z_to_integer => {
        "program t\ncharacter(len=1) :: c = 'Z'\nprint *, transfer(c, 0)\nend program t\n",
        ["90"]
    };

    transfer_char_digit_zero => {
        "program t\ncharacter(len=1) :: c = '0'\nprint *, transfer(c, 0)\nend program t\n",
        ["48"]
    };

    transfer_char_space_to_integer => {
        "program t\ncharacter(len=1) :: c = ' '\nprint *, transfer(c, 0)\nend program t\n",
        ["32"]
    };

    transfer_char_two_bytes_ab => {
        "program t\ncharacter(len=2) :: s = 'AB'\nprint *, transfer(s, 0)\nend program t\n",
        ["16961"]
    };

    transfer_char_roundtrip_single => {
        "program t\ncharacter(len=1) :: c = 'X', d\ninteger :: n\nn = transfer(c, 0)\nd = transfer(n, ' ')\nprint *, ichar(d)\nend program t\n",
        ["88"]
    };

    transfer_char_roundtrip_two_chars => {
        "program t\ncharacter(len=2) :: s = 'Hi', t\ninteger :: n\nn = transfer(s, 0)\nt = transfer(n, '  ')\nprint *, ichar(t(1:1))\nprint *, ichar(t(2:2))\nend program t\n",
        ["72", "105"]
    };

    transfer_int_to_char_ichar_sixty_five => {
        "program t\ninteger :: n = 65\ncharacter(len=1) :: c\n c = transfer(n, ' ')\nprint *, ichar(c)\nend program t\n",
        ["65"]
    };

    transfer_char_four_abcd_roundtrip => {
        "program t\ncharacter(len=4) :: s = 'ABCD', u\ninteger :: n\nn = transfer(s, 0)\nu = transfer(n, '    ')\nprint *, u == s\nend program t\n",
        ["1"]
    };

    transfer_char_array_to_integer_array => {
        "program t\ncharacter(len=1) :: c(3) = ['A', 'B', 'C']\ninteger :: n(3)\nn = transfer(c, n)\nprint *, n(1)\nprint *, n(2)\nprint *, n(3)\nend program t\n",
        ["65", "66", "67"]
    };

    // ── SIZE parameter: truncation and expansion ────────────────────────

    transfer_size_one_from_four_array => {
        "program t\ninteger :: a(4) = [10, 20, 30, 40]\ninteger :: b(1)\nb = transfer(a, b, 1)\nprint *, b(1)\nend program t\n",
        ["10"]
    };

    transfer_size_three_from_four_array => {
        "program t\ninteger :: a(4) = [10, 20, 30, 40]\ninteger :: b(3)\nb = transfer(a, b, 3)\nprint *, b(1)\nprint *, b(2)\nprint *, b(3)\nend program t\n",
        ["10", "20", "30"]
    };

    transfer_size_two_truncated_pair => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: b(2)\nb = transfer(a, b, 2)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["1", "2"]
    };

    transfer_size_expand_scalar_first_element => {
        "program t\ninteger :: a = 42\ninteger :: b(4)\nb = transfer(a, b, 4)\nprint *, b(1)\nend program t\n",
        ["42"]
    };

    transfer_size_two_from_scalar => {
        "program t\ninteger :: a = 99\ninteger :: b(2)\nb = transfer(a, b, 2)\nprint *, b(1)\nend program t\n",
        ["99"]
    };

    transfer_size_on_character_string => {
        "program t\ncharacter(len=4) :: s = 'WXYZ'\ncharacter(len=2) :: t\nt = transfer(s, t, 2)\nprint *, ichar(t(1:1))\nprint *, ichar(t(2:2))\nend program t\n",
        ["87", "88"]
    };

    transfer_size_truncate_byte_array => {
        "program t\ninteger :: n = 305419896\ninteger(kind=1) :: full(4), part(2)\nfull = transfer(n, full)\npart = transfer(full, part, 2)\nprint *, int(part(1))\nprint *, int(part(2))\nend program t\n",
        ["120", "86"]
    };

    transfer_size_larger_than_source_scalar => {
        "program t\ninteger :: a = 7\ninteger :: b(8)\nb = transfer(a, b, 8)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["7", "0"]
    };

    transfer_size_one_real_element => {
        "program t\nreal :: a(3) = [1.0, 2.0, 3.0]\nreal :: b(1)\nb = transfer(a, b, 1)\nprint *, b(1)\nend program t\n",
        ["1"]
    };

    transfer_size_partial_char_to_int => {
        "program t\ncharacter(len=3) :: s = 'abc'\ninteger :: n\nn = transfer(s, 0, 1)\nprint *, n\nend program t\n",
        ["97"]
    };

    transfer_size_expand_kind1_scalar => {
        "program t\ninteger(kind=1) :: source = 42_1\ninteger :: target(2)\ntarget = transfer(source, target, 2)\nprint *, target(1)\nprint *, target(2)\nend program t\n",
        ["42", "0"]
    };

    transfer_logical_scalar_to_integer => {
        "program t\ninteger :: true_bits\ninteger :: false_bits\ntrue_bits = transfer(.true., 0)\nfalse_bits = transfer(.false., 0)\nprint *, true_bits\nprint *, false_bits\nend program t\n",
        ["1", "0"]
    };

    transfer_char_kind1_to_integer_kind2 => {
        "program t\ninteger(kind=2) :: n\nn = transfer('Q', 0_2)\nprint *, n\nend program t\n",
        ["81"]
    };

    // ── Array-level transfers ───────────────────────────────────────────

    transfer_array_int_two_elements => {
        "program t\ninteger :: a(2) = [10, 20]\ninteger :: b(2)\nb = transfer(a, b)\nprint *, b(1)\nprint *, b(2)\nend program t\n",
        ["10", "20"]
    };

    transfer_array_int_four_elements => {
        "program t\ninteger :: a(4) = [1, 2, 3, 4]\ninteger :: b(4)\nb = transfer(a, b)\nprint *, b(1)\nprint *, b(4)\nend program t\n",
        ["1", "4"]
    };

    transfer_array_real_copy => {
        "program t\nreal :: a(3) = [1.5, 2.5, 3.5]\nreal :: b(3)\nb = transfer(a, b)\nprint *, b(2)\nend program t\n",
        ["2.5"]
    };

    transfer_array_byte_from_integer_scalar => {
        "program t\ninteger :: n = 305419896\ninteger(kind=1) :: b(4)\nb = transfer(n, b)\nprint *, int(b(1))\nprint *, int(b(2))\nprint *, int(b(3))\nprint *, int(b(4))\nend program t\n",
        ["120", "86", "52", "18"]
    };

    transfer_array_bytes_to_integer_roundtrip => {
        "program t\ninteger :: original = 305419896\ninteger(kind=1) :: bytes(4)\ninteger :: recovered\nbytes = transfer(original, bytes)\nrecovered = transfer(bytes, 0)\nprint *, recovered\nend program t\n",
        ["305419896"]
    };

    transfer_array_kind8_to_kind4_pair => {
        "program t\ninteger(kind=8) :: big = 1000000000000_8\ninteger(kind=4) :: parts(2)\nparts = transfer(big, parts)\nprint *, transfer(parts, 0_8)\nend program t\n",
        ["1000000000000"]
    };

    transfer_array_character_three_chars => {
        "program t\ncharacter(len=1) :: c(3) = ['f', 'o', 'r']\ncharacter(len=1) :: d(3)\nd = transfer(c, d)\nprint *, ichar(d(1))\nprint *, ichar(d(2))\nprint *, ichar(d(3))\nend program t\n",
        ["102", "111", "114"]
    };

    transfer_array_real_pair_to_scalar_bits => {
        "program t\nreal :: pair(2) = [1.0, 2.0]\ninteger :: n\nn = transfer(pair, 0)\nprint *, transfer(n, pair(1)) == 1.0\nend program t\n",
        ["1"]
    };

    transfer_array_logical_to_integer => {
        "program t\nlogical :: m(2) = [.true., .false.]\ninteger :: n(2)\nn = transfer(m, n)\nprint *, n(1)\nprint *, n(2)\nend program t\n",
        ["1", "0"]
    };

    transfer_array_2d_flatten_to_1d => {
        "program t\ninteger :: m(2,2) = reshape([1, 2, 3, 4], [2, 2])\ninteger :: v(4)\nv = transfer(m, v)\nprint *, v(1)\nprint *, v(4)\nend program t\n",
        ["1", "4"]
    };

    transfer_array_int_in_expression => {
        "program t\ninteger :: a(2) = [5, 6]\ninteger :: b(2)\nb = transfer(a, b)\nprint *, b(1) + b(2)\nend program t\n",
        ["11"]
    };

    transfer_array_kind1_four_bytes => {
        "program t\ninteger(kind=1) :: b(4) = [18_1, 52_1, 86_1, 120_1]\ninteger :: n\nn = transfer(b, 0)\nprint *, n\nend program t\n",
        ["305419896"]
    };
}
