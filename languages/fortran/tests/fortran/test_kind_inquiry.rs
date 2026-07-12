//! Fortran kind/type inquiry intrinsics: kind, selected_*_kind, bit_size,
//! storage_size, precision, range, and digits.

fortran_cases! {
    // ── kind() on literals ────────────────────────────────────────
    kind_integer_literal_one => {
        "program t\nprint *, kind(1)\nend program t\n",
        ["8"]
    };

    kind_real_literal_one_point_zero => {
        "program t\nprint *, kind(1.0)\nend program t\n",
        ["8"]
    };

    kind_double_literal_one_point_zero_d0 => {
        "program t\nprint *, kind(1.0d0)\nend program t\n",
        ["8"]
    };

    kind_logical_literal_true => {
        "program t\nprint *, kind(.true.)\nend program t\n",
        ["8"]
    };

    kind_character_literal_a => {
        "program t\nprint *, kind('a')\nend program t\n",
        ["8"]
    };

    // ── selected_int_kind ─────────────────────────────────────────
    selected_int_kind_range_four => {
        "program t\nprint *, selected_int_kind(4)\nend program t\n",
        ["8"]
    };

    selected_int_kind_range_eight => {
        "program t\nprint *, selected_int_kind(8)\nend program t\n",
        ["8"]
    };

    selected_int_kind_range_sixteen => {
        "program t\nprint *, selected_int_kind(16)\nend program t\n",
        ["8"]
    };

    selected_int_kind_range_thirty_two_unavailable => {
        "program t\nprint *, selected_int_kind(32)\nend program t\n",
        ["8"]
    };

    selected_int_kind_range_two => {
        "program t\nprint *, selected_int_kind(2)\nend program t\n",
        ["8"]
    };

    // ── selected_real_kind ────────────────────────────────────────
    selected_real_kind_precision_six => {
        "program t\nprint *, selected_real_kind(6)\nend program t\n",
        ["8"]
    };

    selected_real_kind_precision_fifteen => {
        "program t\nprint *, selected_real_kind(15)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p15_r307 => {
        "program t\nprint *, selected_real_kind(15, 307)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p6_r37 => {
        "program t\nprint *, selected_real_kind(6, 37)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p6_r99 => {
        "program t\nprint *, selected_real_kind(6, 99)\nend program t\n",
        ["8"]
    };

    selected_real_kind_p7_r37 => {
        "program t\nprint *, selected_real_kind(7, 37)\nend program t\n",
        ["8"]
    };

    // ── kind on typed variables ───────────────────────────────────
    kind_default_integer_variable => {
        "program t\ninteger :: n = 7\nprint *, kind(n)\nend program t\n",
        ["8"]
    };

    kind_int_kind_eight_variable => {
        "program t\ninteger(kind=8) :: n = 7_8\nprint *, kind(n)\nend program t\n",
        ["8"]
    };

    kind_default_real_variable => {
        "program t\nreal :: x = 1.5\nprint *, kind(x)\nend program t\n",
        ["8"]
    };

    kind_real_kind_eight_variable => {
        "program t\nreal(kind=8) :: x = 1.5_8\nprint *, kind(x)\nend program t\n",
        ["8"]
    };

    kind_of_kind_result_is_eight => {
        "program t\nprint *, kind(kind(1))\nend program t\n",
        ["8"]
    };

    // ── bit_size on integers ──────────────────────────────────────
    bit_size_default_integer_literal => {
        "program t\nprint *, bit_size(0)\nend program t\n",
        ["32"]
    };

    bit_size_default_integer_variable => {
        "program t\ninteger :: x = 0\nprint *, bit_size(x)\nend program t\n",
        ["32"]
    };

    bit_size_int_kind_eight_variable => {
        "program t\ninteger(kind=8) :: x = 0_8\nprint *, bit_size(x)\nend program t\n",
        ["64"]
    };

    bit_size_int_kind_two_variable => {
        "program t\ninteger(kind=2) :: x = 0_2\nprint *, bit_size(x)\nend program t\n",
        ["16"]
    };

    bit_size_int_kind_one_variable => {
        "program t\ninteger(kind=1) :: x = 0_1\nprint *, bit_size(x)\nend program t\n",
        ["8"]
    };

    // ── bit_size on reals ─────────────────────────────────────────
    bit_size_default_real_literal => {
        "program t\nprint *, bit_size(0.0)\nend program t\n",
        ["64"]
    };

    bit_size_default_real_variable => {
        "program t\nreal :: x = 0.0\nprint *, bit_size(x)\nend program t\n",
        ["64"]
    };

    bit_size_real_kind_eight_variable => {
        "program t\nreal(kind=8) :: x = 0.0_8\nprint *, bit_size(x)\nend program t\n",
        ["64"]
    };

    bit_size_real_kind_four_variable => {
        "program t\nreal(kind=4) :: x = 0.0_4\nprint *, bit_size(x)\nend program t\n",
        ["32"]
    };

    // ── storage_size on integers and reals ────────────────────────
    storage_size_default_integer_variable => {
        "program t\ninteger :: x = 0\nprint *, storage_size(x)\nend program t\n",
        ["32"]
    };

    storage_size_int_kind_eight_variable => {
        "program t\ninteger(kind=8) :: x = 0_8\nprint *, storage_size(x)\nend program t\n",
        ["64"]
    };

    storage_size_default_real_variable => {
        "program t\nreal :: x = 0.0\nprint *, storage_size(x)\nend program t\n",
        ["64"]
    };

    storage_size_real_kind_eight_variable => {
        "program t\nreal(kind=8) :: x = 0.0_8\nprint *, storage_size(x)\nend program t\n",
        ["64"]
    };

    bit_size_equals_storage_size_default_integer => {
        "program t\ninteger :: x = 0\nprint *, bit_size(x) == storage_size(x)\nend program t\n",
        ["true"]
    };

    // ── precision on real types ───────────────────────────────────
    precision_default_real_variable => {
        "program t\nreal :: x = 0.0\nprint *, precision(x)\nend program t\n",
        ["53"]
    };

    precision_real_kind_eight_variable => {
        "program t\nreal(kind=8) :: x = 0.0_8\nprint *, precision(x)\nend program t\n",
        ["53"]
    };

    precision_real_kind_four_variable => {
        "program t\nreal(kind=4) :: x = 0.0_4\nprint *, precision(x)\nend program t\n",
        ["24"]
    };

    precision_default_real_literal => {
        "program t\nprint *, precision(0.0)\nend program t\n",
        ["53"]
    };

    // ── range on integer and real types ───────────────────────────
    range_default_integer_variable => {
        "program t\ninteger :: x = 0\nprint *, range(x)\nend program t\n",
        ["9"]
    };

    range_int_kind_eight_variable => {
        "program t\ninteger(kind=8) :: x = 0_8\nprint *, range(x)\nend program t\n",
        ["18"]
    };

    range_int_kind_two_variable => {
        "program t\ninteger(kind=2) :: x = 0_2\nprint *, range(x)\nend program t\n",
        ["4"]
    };

    range_default_real_variable => {
        "program t\nreal :: x = 0.0\nprint *, range(x)\nend program t\n",
        ["307"]
    };

    range_real_kind_eight_variable => {
        "program t\nreal(kind=8) :: x = 0.0_8\nprint *, range(x)\nend program t\n",
        ["307"]
    };

    range_real_kind_four_variable => {
        "program t\nreal(kind=4) :: x = 0.0_4\nprint *, range(x)\nend program t\n",
        ["37"]
    };

    // ── digits inquiry ────────────────────────────────────────────
    digits_default_integer_variable => {
        "program t\ninteger :: x = 0\nprint *, digits(x)\nend program t\n",
        ["9"]
    };

    digits_int_kind_eight_variable => {
        "program t\ninteger(kind=8) :: x = 0_8\nprint *, digits(x)\nend program t\n",
        ["18"]
    };

    digits_default_real_variable => {
        "program t\nreal :: x = 0.0\nprint *, digits(x)\nend program t\n",
        ["15"]
    };

    digits_real_kind_four_variable => {
        "program t\nreal(kind=4) :: x = 0.0_4\nprint *, digits(x)\nend program t\n",
        ["6"]
    };

    // ── cross-kind consistency checks ─────────────────────────────
    selected_int_kind_matches_kind_of_literal => {
        "program t\ninteger, parameter :: k = selected_int_kind(9)\nprint *, k\nend program t\n",
        ["8"]
    };

    selected_real_kind_matches_kind_of_double_literal => {
        "program t\ninteger, parameter :: k = selected_real_kind(15, 307)\nprint *, k\nend program t\n",
        ["8"]
    };

    range_integer_less_than_int_kind_eight => {
        "program t\ninteger :: s = 0\ninteger(kind=8) :: b = 0_8\nprint *, range(s) < range(b)\nend program t\n",
        ["true"]
    };

    precision_real_kind_four_less_than_default => {
        "program t\nreal(kind=4) :: s = 0.0_4\nreal :: d = 0.0\nprint *, precision(s) < precision(d)\nend program t\n",
        ["true"]
    };

    bit_size_logical_literal => {
        "program t\nprint *, bit_size(.true.)\nend program t\n",
        ["32"]
    };
}
