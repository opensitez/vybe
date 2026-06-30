//! Extended ENUM types: ENUM/BIND(C) definitions, auto/explicit values,
//! enum use in expressions, SELECT CASE, arrays, modules, and I/O.
//! Distinct from `test_fortran2003.rs` (three basic enum compile tests).

fortran_cases! {
    // ── ENUM BIND(C) basic definitions ───────────────────────────────

    enum_bind_c_rgb_values => {
        "program t\nenum, bind(c)\nenumerator :: RED = 0, GREEN = 1, BLUE = 2\nend enum\ninteger :: c = GREEN\nprint *, c\nend program t\n",
        ["1"]
    };

    enum_bind_c_cardinal_directions => {
        "program t\nenum, bind(c)\nenumerator :: NORTH, SOUTH, EAST, WEST\nend enum\ninteger :: d = EAST\nprint *, d\nend program t\n",
        ["2"]
    };

    enum_bind_c_priority_levels => {
        "program t\nenum, bind(c)\nenumerator :: LOW = 1, MEDIUM = 5, HIGH = 10\nend enum\ninteger :: p = HIGH\nprint *, p\nend program t\n",
        ["10"]
    };

    enum_bind_c_single_member => {
        "program t\nenum, bind(c)\nenumerator :: ONLY = 42\nend enum\ninteger :: v = ONLY\nprint *, v\nend program t\n",
        ["42"]
    };

    enum_bind_c_zero_start => {
        "program t\nenum, bind(c)\nenumerator :: FIRST = 0, SECOND, THIRD\nend enum\ninteger :: v = THIRD\nprint *, v\nend program t\n",
        ["2"]
    };

    enum_bind_c_negative_start => {
        "program t\nenum, bind(c)\nenumerator :: MINUS2 = -2, MINUS1, ZERO\nend enum\ninteger :: v = ZERO\nprint *, v\nend program t\n",
        ["0"]
    };

    enum_bind_c_large_gap => {
        "program t\nenum, bind(c)\nenumerator :: A = 100, B = 200, C = 300\nend enum\ninteger :: v = B\nprint *, v\nend program t\n",
        ["200"]
    };

    enum_bind_c_hex_style_values => {
        "program t\nenum, bind(c)\nenumerator :: FLAG_A = 1, FLAG_B = 2, FLAG_C = 4, FLAG_D = 8\nend enum\ninteger :: f = FLAG_C\nprint *, f\nend program t\n",
        ["4"]
    };

    // ── Auto-increment chains ──────────────────────────────────────────

    enum_auto_four_items => {
        "program t\nenum, bind(c)\nenumerator :: A, B, C, D\nend enum\nprint *, A\nprint *, D\nend program t\n",
        ["0", "3"]
    };

    enum_auto_after_explicit_start => {
        "program t\nenum, bind(c)\nenumerator :: START = 10, NEXT, LAST\nend enum\nprint *, START\nprint *, NEXT\nprint *, LAST\nend program t\n",
        ["10", "11", "12"]
    };

    enum_auto_ten_members => {
        "program t\nenum, bind(c)\nenumerator :: E0, E1, E2, E3, E4, E5, E6, E7, E8, E9\nend enum\nprint *, E0\nprint *, E9\nend program t\n",
        ["0", "9"]
    };

    enum_auto_mid_chain_restart => {
        "program t\nenum, bind(c)\nenumerator :: X = 5, Y, Z = 20, W\nend enum\nprint *, Y\nprint *, W\nend program t\n",
        ["6", "21"]
    };

    // ── Enum values in expressions ───────────────────────────────────

    enum_expr_addition => {
        "program t\nenum, bind(c)\nenumerator :: A = 1, B = 2, C = 3\nend enum\nprint *, A + B\nend program t\n",
        ["3"]
    };

    enum_expr_subtraction => {
        "program t\nenum, bind(c)\nenumerator :: A = 10, B = 3\nend enum\nprint *, A - B\nend program t\n",
        ["7"]
    };

    enum_expr_multiplication => {
        "program t\nenum, bind(c)\nenumerator :: A = 4, B = 5\nend enum\nprint *, A * B\nend program t\n",
        ["20"]
    };

    enum_expr_division => {
        "program t\nenum, bind(c)\nenumerator :: A = 20, B = 4\nend enum\nprint *, A / B\nend program t\n",
        ["5"]
    };

    enum_expr_mod => {
        "program t\nenum, bind(c)\nenumerator :: A = 17, B = 5\nend enum\nprint *, mod(A, B)\nend program t\n",
        ["2"]
    };

    enum_expr_comparison_equal => {
        "program t\nenum, bind(c)\nenumerator :: A = 5, B = 5\nend enum\nprint *, A == B\nend program t\n",
        ["true"]
    };

    enum_expr_comparison_less => {
        "program t\nenum, bind(c)\nenumerator :: A = 1, B = 9\nend enum\nprint *, A < B\nend program t\n",
        ["true"]
    };

    enum_expr_comparison_greater => {
        "program t\nenum, bind(c)\nenumerator :: A = 9, B = 1\nend enum\nprint *, A > B\nend program t\n",
        ["true"]
    };

    enum_expr_max_of_two => {
        "program t\nenum, bind(c)\nenumerator :: LOW = 2, HIGH = 8\nend enum\nprint *, max(LOW, HIGH)\nend program t\n",
        ["8"]
    };

    enum_expr_min_of_two => {
        "program t\nenum, bind(c)\nenumerator :: LOW = 2, HIGH = 8\nend enum\nprint *, min(LOW, HIGH)\nend program t\n",
        ["2"]
    };

    enum_expr_abs_negative_member => {
        "program t\nenum, bind(c)\nenumerator :: NEG = -7\nend enum\nprint *, abs(NEG)\nend program t\n",
        ["7"]
    };

    // ── Enum in SELECT CASE ────────────────────────────────────────────

    enum_select_case_first => {
        "program t\nenum, bind(c)\nenumerator :: ONE = 1, TWO = 2, THREE = 3\nend enum\ninteger :: v = ONE\nselect case (v)\ncase (ONE)\nprint *, 10\ncase (TWO)\nprint *, 20\ncase default\nprint *, 0\nend select\nend program t\n",
        ["10"]
    };

    enum_select_case_second => {
        "program t\nenum, bind(c)\nenumerator :: ONE = 1, TWO = 2, THREE = 3\nend enum\ninteger :: v = TWO\nselect case (v)\ncase (ONE)\nprint *, 10\ncase (TWO)\nprint *, 20\ncase default\nprint *, 0\nend select\nend program t\n",
        ["20"]
    };

    enum_select_case_default => {
        "program t\nenum, bind(c)\nenumerator :: ONE = 1, TWO = 2\nend enum\ninteger :: v = 99\nselect case (v)\ncase (ONE)\nprint *, 1\ncase (TWO)\nprint *, 2\ncase default\nprint *, 99\nend select\nend program t\n",
        ["99"]
    };

    enum_select_case_range => {
        "program t\nenum, bind(c)\nenumerator :: LOW = 1, MID = 5, HIGH = 10\nend enum\ninteger :: v = MID\nselect case (v)\ncase (1:3)\nprint *, 1\ncase (4:6)\nprint *, 2\ncase default\nprint *, 3\nend select\nend program t\n",
        ["2"]
    };

    enum_select_case_cardinal => {
        "program t\nenum, bind(c)\nenumerator :: N, S, E, W\nend enum\ninteger :: d = W\nselect case (d)\ncase (N)\nprint *, 1\ncase (S)\nprint *, 2\ncase (E)\nprint *, 3\ncase (W)\nprint *, 4\nend select\nend program t\n",
        ["4"]
    };

    // ── Enum in IF constructs ──────────────────────────────────────────

    enum_if_equals_member => {
        "program t\nenum, bind(c)\nenumerator :: ON = 1, OFF = 0\nend enum\ninteger :: s = ON\nif (s == ON) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };

    enum_if_not_equal => {
        "program t\nenum, bind(c)\nenumerator :: ON = 1, OFF = 0\nend enum\ninteger :: s = OFF\nif (s /= ON) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };

    enum_if_greater_chain => {
        "program t\nenum, bind(c)\nenumerator :: LOW = 1, MID = 5, HIGH = 10\nend enum\ninteger :: v = MID\nif (v > LOW) then\nif (v < HIGH) then\nprint *, 1\nelse\nprint *, 0\nend if\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };

    // ── Enum in arrays ─────────────────────────────────────────────────

    enum_array_indexed_store => {
        "program t\nenum, bind(c)\nenumerator :: A = 1, B = 2, C = 3\nend enum\ninteger :: arr(3)\narr(A) = 10\narr(B) = 20\narr(C) = 30\nprint *, arr(B)\nend program t\n",
        ["20"]
    };

    enum_array_initializer => {
        "program t\nenum, bind(c)\nenumerator :: R = 0, G = 1, B = 2\nend enum\ninteger :: codes(3) = [R, G, B]\nprint *, codes(2)\nend program t\n",
        ["1"]
    };

    enum_array_sum_members => {
        "program t\nenum, bind(c)\nenumerator :: A = 1, B = 2, C = 3, D = 4\nend enum\ninteger :: vals(4) = [A, B, C, D]\nprint *, sum(vals)\nend program t\n",
        ["10"]
    };

    enum_array_lookup => {
        "program t\nenum, bind(c)\nenumerator :: RED = 0, GREEN = 1, BLUE = 2\nend enum\ncharacter(len=5) :: names(0:2)\nnames(RED) = 'red'\nnames(GREEN) = 'grn'\nnames(BLUE) = 'blu'\nprint *, trim(names(GREEN))\nend program t\n",
        ["grn"]
    };

    // ── Enum in DO loops ───────────────────────────────────────────────

    enum_do_loop_bound => {
        "program t\nenum, bind(c)\nenumerator :: START = 1, LAST = 5\nend enum\ninteger :: i, s\ns = 0\ndo i = START, LAST\ns = s + i\nend do\nprint *, s\nend program t\n",
        ["15"]
    };

    enum_do_loop_step => {
        "program t\nenum, bind(c)\nenumerator :: BEGIN = 0, STEP = 2, LIMIT = 10\nend enum\ninteger :: i, c\nc = 0\ndo i = BEGIN, LIMIT, STEP\nc = c + 1\nend do\nprint *, c\nend program t\n",
        ["6"]
    };

    enum_do_concurrent_with_enum => {
        "program t\nenum, bind(c)\nenumerator :: N = 5\nend enum\ninteger :: a(N)\ndo concurrent (i = 1:N)\na(i) = i\nend do\nprint *, a(N)\nend program t\n",
        ["5"]
    };

    // ── Enum in derived types ──────────────────────────────────────────

    enum_dtype_field => {
        "program t\nenum, bind(c)\nenumerator :: IDLE = 0, RUN = 1, DONE = 2\nend enum\ntype :: Task\ninteger :: state\nend type Task\ntype(Task) :: t\nt%state = RUN\nprint *, t%state\nend program t\n",
        ["1"]
    };

    enum_dtype_array_field => {
        "program t\nenum, bind(c)\nenumerator :: A = 0, B = 1, C = 2\nend enum\ntype :: Set\ninteger :: tags(3)\nend type Set\ntype(Set) :: s\ns%tags = [A, B, C]\nprint *, s%tags(B + 1)\nend program t\n",
        ["1"]
    };

    // ── Module-scoped enums ────────────────────────────────────────────

    enum_module_export => {
        "module colors\nenum, bind(c)\nenumerator :: RED = 0, GREEN = 1, BLUE = 2\nend enum\nend module colors\nprogram t\nuse colors\nprint *, GREEN\nend program t\n",
        ["1"]
    };

    enum_module_with_subroutine => {
        "module status\nenum, bind(c)\nenumerator :: OK = 0, ERR = 1\nend enum\ncontains\ninteger function code() result(c)\nc = OK\nend function code\nend module status\nprogram t\nuse status\nprint *, code()\nend program t\n",
        ["0"]
    };

    enum_module_two_sets => {
        "module dirs\nenum, bind(c)\nenumerator :: N = 0, S = 1\nend enum\nenum, bind(c)\nenumerator :: E = 0, W = 1\nend enum\nend module dirs\nprogram t\nuse dirs\nprint *, N + E\nend program t\n",
        ["0"]
    };

    // ── Enum I/O (print values) ────────────────────────────────────────

    enum_io_print_single => {
        "program t\nenum, bind(c)\nenumerator :: VAL = 42\nend enum\nprint *, VAL\nend program t\n",
        ["42"]
    };

    enum_io_print_multiple => {
        "program t\nenum, bind(c)\nenumerator :: A = 1, B = 2, C = 3\nend enum\nprint *, A\nprint *, B\nprint *, C\nend program t\n",
        ["1", "2", "3"]
    };

    enum_io_print_auto_chain => {
        "program t\nenum, bind(c)\nenumerator :: X, Y, Z\nend enum\nprint *, X\nprint *, Y\nprint *, Z\nend program t\n",
        ["0", "1", "2"]
    };

    enum_io_print_expression => {
        "program t\nenum, bind(c)\nenumerator :: BASE = 10, OFFSET = 3\nend enum\nprint *, BASE + OFFSET\nend program t\n",
        ["13"]
    };

    enum_io_print_in_loop => {
        "program t\nenum, bind(c)\nenumerator :: V0 = 0, V1 = 1, V2 = 2\nend enum\ninteger :: vals(3) = [V0, V1, V2]\ninteger :: i\ndo i = 1, 3\nprint *, vals(i)\nend do\nend program t\n",
        ["0", "1", "2"]
    };

    // ── Multiple enums in one program ──────────────────────────────────

    enum_two_independent_sets => {
        "program t\nenum, bind(c)\nenumerator :: RED = 0, GREEN = 1\nend enum\nenum, bind(c)\nenumerator :: UP = 0, DOWN = 1\nend enum\nprint *, RED + DOWN\nend program t\n",
        ["1"]
    };

    enum_assign_from_member => {
        "program t\nenum, bind(c)\nenumerator :: A = 5, B = 10\nend enum\ninteger :: x\nx = A\nprint *, x\nx = B\nprint *, x\nend program t\n",
        ["5", "10"]
    };

    enum_switch_via_assignment => {
        "program t\nenum, bind(c)\nenumerator :: MODE_A = 1, MODE_B = 2\nend enum\ninteger :: mode\nmode = MODE_A\nprint *, mode\nmode = MODE_B\nprint *, mode\nend program t\n",
        ["1", "2"]
    };

    // ── Edge cases ─────────────────────────────────────────────────────

    enum_member_as_array_size => {
        "program t\nenum, bind(c)\nenumerator :: SIZE = 4\nend enum\ninteger :: arr(SIZE)\narr = [1, 2, 3, 4]\nprint *, arr(SIZE)\nend program t\n",
        ["4"]
    };

    enum_member_in_parameter => {
        "program t\nenum, bind(c)\nenumerator :: MAX = 100\nend enum\ninteger, parameter :: limit = MAX\nprint *, limit\nend program t\n",
        ["100"]
    };

    enum_equality_all_members => {
        "program t\nenum, bind(c)\nenumerator :: X = 7\nend enum\nif (X == 7) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };

    enum_ior_combine_flags => {
        "program t\nenum, bind(c)\nenumerator :: F1 = 1, F2 = 2, F4 = 4\nend enum\nprint *, ior(F1, F2)\nend program t\n",
        ["3"]
    };

    enum_ieor_toggle => {
        "program t\nenum, bind(c)\nenumerator :: A = 5, B = 3\nend enum\nprint *, ieor(A, B)\nend program t\n",
        ["6"]
    };

    enum_iand_mask => {
        "program t\nenum, bind(c)\nenumerator :: MASK = 7, VAL = 5\nend enum\nprint *, iand(MASK, VAL)\nend program t\n",
        ["5"]
    };

    enum_merge_select => {
        "program t\nenum, bind(c)\nenumerator :: CHOICE_A = 10, CHOICE_B = 20\nend enum\nprint *, merge(CHOICE_A, CHOICE_B, .true.)\nend program t\n",
        ["10"]
    };

    enum_nested_select_in_function => {
        "program t\nprint *, label(2)\ncontains\ninteger function label(v) result(r)\nenum, bind(c)\nenumerator :: ONE = 1, TWO = 2\nend enum\nselect case (v)\ncase (ONE)\nr = 10\ncase (TWO)\nr = 20\ncase default\nr = 0\nend select\nend function label\nend program t\n",
        ["20"]
    };

    enum_contains_subroutine => {
        "program t\ncall show()\ncontains\nsubroutine show()\nenum, bind(c)\nenumerator :: TAG = 99\nend enum\nprint *, TAG\nend subroutine show\nend program t\n",
        ["99"]
    };

    enum_sign_function => {
        "program t\nenum, bind(c)\nenumerator :: POS = 5, NEG = -5\nend enum\nprint *, sign(POS, NEG)\nend program t\n",
        ["-5"]
    };
}
