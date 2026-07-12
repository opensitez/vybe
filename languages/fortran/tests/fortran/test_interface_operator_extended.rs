//! Extended INTERFACE operator and generic procedure coverage: derived-type
//! arithmetic, assignment overloads, module procedure bindings, and multi-signature
//! generics. Distinct from `test_modules_advanced.rs` and operator cases in
//! `test_module_use_extended.rs`.

fortran_cases! {
    // ── operator(+): derived-type addition ───────────────────────────

    operator_plus_point2d_sum_components => {
        "module gpoint\nimplicit none\ntype :: Point\nreal :: x, y\nend type Point\ninterface operator(+)\nmodule procedure add_point\nend interface\ncontains\nfunction add_point(a, b) result(c)\ntype(Point), intent(in) :: a, b\ntype(Point) :: c\nc%x = a%x + b%x\nc%y = a%y + b%y\nend function add_point\nend module gpoint\nprogram t\nuse gpoint\ntype(Point) :: p, q, r\np%x = 1.0; p%y = 2.0\nq%x = 3.0; q%y = 4.0\nr = p + q\nprint *, int(r%x)\nprint *, int(r%y)\nend program t\n",
        ["4", "6"]
    };

    operator_plus_vector1d_lengths => {
        "module gvec\nimplicit none\ntype :: Vec\ninteger :: n\nend type Vec\ninterface operator(+)\nmodule procedure add_vec\nend interface\ncontains\nfunction add_vec(a, b) result(c)\ntype(Vec), intent(in) :: a, b\ntype(Vec) :: c\nc%n = a%n + b%n\nend function add_vec\nend module gvec\nprogram t\nuse gvec\ntype(Vec) :: u, v, w\nu%n = 5\nv%n = 7\nw = u + v\nprint *, w%n\nend program t\n",
        ["12"]
    };

    operator_plus_complex_like_pair => {
        "module gcplx\nimplicit none\ntype :: Cplx\nreal :: re, im\nend type Cplx\ninterface operator(+)\nmodule procedure add_cplx\nend interface\ncontains\nfunction add_cplx(a, b) result(c)\ntype(Cplx), intent(in) :: a, b\ntype(Cplx) :: c\nc%re = a%re + b%re\nc%im = a%im + b%im\nend function add_cplx\nend module gcplx\nprogram t\nuse gcplx\ntype(Cplx) :: a, b, c\na%re = 1.0; a%im = 2.0\nb%re = 3.0; b%im = -1.0\nc = a + b\nprint *, int(c%re)\nprint *, int(c%im)\nend program t\n",
        ["4", "1"]
    };

    operator_plus_integer_wrapper => {
        "module gwrap\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface operator(+)\nmodule procedure add_box\nend interface\ncontains\nfunction add_box(a, b) result(c)\ntype(Box), intent(in) :: a, b\ntype(Box) :: c\nc%v = a%v + b%v\nend function add_box\nend module gwrap\nprogram t\nuse gwrap\ntype(Box) :: x, y, z\nx%v = 10\ny%v = 15\nz = x + y\nprint *, z%v\nend program t\n",
        ["25"]
    };

    operator_plus_mixed_scalar_on_type => {
        "module gshift\nimplicit none\ntype :: Offset\ninteger :: delta\nend type Offset\ninterface operator(+)\nmodule procedure add_offset\nend interface\ncontains\nfunction add_offset(a, b) result(c)\ntype(Offset), intent(in) :: a, b\ntype(Offset) :: c\nc%delta = a%delta + b%delta\nend function add_offset\nend module gshift\nprogram t\nuse gshift\ntype(Offset) :: a, b, c\na%delta = 4\nb%delta = -1\nc = a + b\nprint *, c%delta\nend program t\n",
        ["3"]
    };

    // ── operator(-): unary and binary ────────────────────────────────

    operator_unary_minus_on_signed => {
        "module gneg\nimplicit none\ntype :: Signed\ninteger :: v\nend type Signed\ninterface operator(-)\nmodule procedure negate_signed\nend interface\ncontains\nfunction negate_signed(a) result(b)\ntype(Signed), intent(in) :: a\ntype(Signed) :: b\nb%v = -a%v\nend function negate_signed\nend module gneg\nprogram t\nuse gneg\ntype(Signed) :: x, y\nx%v = 12\ny = -x\nprint *, y%v\nend program t\n",
        ["-12"]
    };

    operator_binary_minus_point_diff => {
        "module gpdiff\nimplicit none\ntype :: Point\ninteger :: x, y\nend type Point\ninterface operator(-)\nmodule procedure sub_point\nend interface\ncontains\nfunction sub_point(a, b) result(c)\ntype(Point), intent(in) :: a, b\ntype(Point) :: c\nc%x = a%x - b%x\nc%y = a%y - b%y\nend function sub_point\nend module gpdiff\nprogram t\nuse gpdiff\ntype(Point) :: a, b, c\na%x = 9; a%y = 7\nb%x = 4; b%y = 2\nc = a - b\nprint *, c%x\nprint *, c%y\nend program t\n",
        ["5", "5"]
    };

    operator_unary_plus_identity_box => {
        "module gplus\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface operator(+)\nmodule procedure id_box\nend interface\ncontains\nfunction id_box(a) result(b)\ntype(Box), intent(in) :: a\ntype(Box) :: b\nb%v = +a%v\nend function id_box\nend module gplus\nprogram t\nuse gplus\ntype(Box) :: x, y\nx%v = 8\ny = +x\nprint *, y%v\nend program t\n",
        ["8"]
    };

    // ── operator(*), operator(/) ─────────────────────────────────────

    operator_multiply_scalar_on_box => {
        "module gmul\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface operator(*)\nmodule procedure mul_box\nend interface\ncontains\nfunction mul_box(a, b) result(c)\ntype(Box), intent(in) :: a, b\ntype(Box) :: c\nc%v = a%v * b%v\nend function mul_box\nend module gmul\nprogram t\nuse gmul\ntype(Box) :: a, b, c\na%v = 6\nb%v = 7\nc = a * b\nprint *, c%v\nend program t\n",
        ["42"]
    };

    operator_divide_ratio_pair => {
        "module gdiv\nimplicit none\ntype :: Ratio\ninteger :: num, den\nend type Ratio\ninterface operator(/)\nmodule procedure div_ratio\nend interface\ncontains\nfunction div_ratio(a, b) result(c)\ntype(Ratio), intent(in) :: a, b\ntype(Ratio) :: c\nc%num = a%num * b%den\nc%den = a%den * b%num\nend function div_ratio\nend module gdiv\nprogram t\nuse gdiv\ntype(Ratio) :: a, b, c\na%num = 1; a%den = 2\nb%num = 3; b%den = 4\nc = a / b\nprint *, c%num\nprint *, c%den\nend program t\n",
        ["4", "6"]
    };

    operator_multiply_repeated_addition => {
        "module gscale\nimplicit none\ntype :: Weight\ninteger :: grams\nend type Weight\ninterface operator(*)\nmodule procedure scale_weight\nend interface\ncontains\nfunction scale_weight(w, n) result(r)\ntype(Weight), intent(in) :: w\ninteger, intent(in) :: n\ntype(Weight) :: r\nr%grams = w%grams * n\nend function scale_weight\nend module gscale\nprogram t\nuse gscale\ntype(Weight) :: w, r\nw%grams = 5\nr = w * 4\nprint *, r%grams\nend program t\n",
        ["20"]
    };

    // ── operator(//): character derived types ────────────────────────

    operator_concat_labels => {
        "module glabel\nimplicit none\ntype :: Label\ncharacter(len=8) :: text\nend type Label\ninterface operator(//)\nmodule procedure concat_label\nend interface\ncontains\nfunction concat_label(a, b) result(c)\ntype(Label), intent(in) :: a, b\ntype(Label) :: c\nc%text = trim(a%text) // trim(b%text)\nend function concat_label\nend module glabel\nprogram t\nuse glabel\ntype(Label) :: a, b, c\na%text = 'foo'\nb%text = 'bar'\nc = a // b\nprint *, trim(c%text)\nend program t\n",
        ["foobar"]
    };

    operator_concat_with_space_join => {
        "module gjoin\nimplicit none\ntype :: Token\ncharacter(len=6) :: word\nend type Token\ninterface operator(//)\nmodule procedure join_token\nend interface\ncontains\nfunction join_token(a, b) result(c)\ntype(Token), intent(in) :: a, b\ntype(Token) :: c\nc%word = trim(a%word) // '-' // trim(b%word)\nend function join_token\nend module gjoin\nprogram t\nuse gjoin\ntype(Token) :: x, y, z\nx%word = 'ab'\ny%word = 'cd'\nz = x // y\nprint *, trim(z%word)\nend program t\n",
        ["ab-cd"]
    };

    // ── assignment(=): typed assignment overloads ────────────────────

    assignment_int_to_box => {
        "module gassign\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface assignment(=)\nmodule procedure int_to_box\nend interface\ncontains\nsubroutine int_to_box(dest, src)\ntype(Box), intent(out) :: dest\ninteger, intent(in) :: src\ndest%v = src\nend subroutine int_to_box\nend module gassign\nprogram t\nuse gassign\ntype(Box) :: b\nb = 42\nprint *, b%v\nend program t\n",
        ["42"]
    };

    assignment_box_to_box_copy => {
        "module gcopy\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface assignment(=)\nmodule procedure copy_box\nend interface\ncontains\nsubroutine copy_box(dest, src)\ntype(Box), intent(out) :: dest\ntype(Box), intent(in) :: src\ndest%v = src%v + 1\nend subroutine copy_box\nend module gcopy\nprogram t\nuse gcopy\ntype(Box) :: a, b\na%v = 10\nb = a\nprint *, b%v\nend program t\n",
        ["11"]
    };

    assignment_real_to_metric => {
        "module gmetric\nimplicit none\ntype :: Metric\nreal :: value\nend type Metric\ninterface assignment(=)\nmodule procedure real_to_metric\nend interface\ncontains\nsubroutine real_to_metric(dest, src)\ntype(Metric), intent(out) :: dest\nreal, intent(in) :: src\ndest%value = src * 2.0\nend subroutine real_to_metric\nend module gmetric\nprogram t\nuse gmetric\ntype(Metric) :: m\nm = 3.0\nprint *, int(m%value)\nend program t\n",
        ["6"]
    };

    assignment_character_to_name => {
        "module gname\nimplicit none\ntype :: Name\ncharacter(len=10) :: text\nend type Name\ninterface assignment(=)\nmodule procedure char_to_name\nend interface\ncontains\nsubroutine char_to_name(dest, src)\ntype(Name), intent(out) :: dest\ncharacter(len=*), intent(in) :: src\ndest%text = src\nend subroutine char_to_name\nend module gname\nprogram t\nuse gname\ntype(Name) :: n\nn = 'Fortran'\nprint *, trim(n%text)\nend program t\n",
        ["Fortran"]
    };

    // ── operator(==) and relational overloads ────────────────────────

    operator_eq_boxes_compare_value => {
        "module geq\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface operator(==)\nmodule procedure eq_box\nend interface\ncontains\nfunction eq_box(a, b) result(r)\ntype(Box), intent(in) :: a, b\nlogical :: r\nr = a%v == b%v\nend function eq_box\nend module geq\nprogram t\nuse geq\ntype(Box) :: a, b\na%v = 5\nb%v = 5\nprint *, a == b\nend program t\n",
        ["true"]
    };

    operator_eq_boxes_not_equal => {
        "module geq2\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface operator(==)\nmodule procedure eq_box\nend interface\ncontains\nfunction eq_box(a, b) result(r)\ntype(Box), intent(in) :: a, b\nlogical :: r\nr = a%v == b%v\nend function eq_box\nend module geq2\nprogram t\nuse geq2\ntype(Box) :: a, b\na%v = 5\nb%v = 6\nprint *, a == b\nend program t\n",
        ["false"]
    };

    operator_lt_ordered_pair => {
        "module glt\nimplicit none\ntype :: Pair\ninteger :: a, b\nend type Pair\ninterface operator(<)\nmodule procedure lt_pair\nend interface\ncontains\nfunction lt_pair(x, y) result(r)\ntype(Pair), intent(in) :: x, y\nlogical :: r\nr = x%a < y%a\nend function lt_pair\nend module glt\nprogram t\nuse glt\ntype(Pair) :: p, q\np%a = 2\nq%a = 5\nprint *, p < q\nend program t\n",
        ["true"]
    };

    // ── Generic interface: module procedure lists ────────────────────

    generic_add_int_and_real => {
        "module gadd\nimplicit none\ninterface add_generic\nmodule procedure add_int, add_real\nend interface\ncontains\nfunction add_int(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a + b\nend function add_int\nfunction add_real(a, b) result(r)\nreal, intent(in) :: a, b\nreal :: r\nr = a + b\nend function add_real\nend module gadd\nprogram t\nuse gadd\nprint *, add_generic(2, 3)\nprint *, int(add_generic(1.5, 2.5))\nend program t\n",
        ["5", "4"]
    };

    generic_max_three_kinds => {
        "module gmax\nimplicit none\ninterface pick_max\nmodule procedure max_int, max_real, max_logical\nend interface\ncontains\nfunction max_int(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = max(a, b)\nend function max_int\nfunction max_real(a, b) result(r)\nreal, intent(in) :: a, b\nreal :: r\nr = max(a, b)\nend function max_real\nfunction max_logical(a, b) result(r)\nlogical, intent(in) :: a, b\nlogical :: r\nif (a .eqv. b) then\nr = a\nelse\nr = .true.\nend if\nend function max_logical\nend module gmax\nprogram t\nuse gmax\nprint *, pick_max(2, 9)\nprint *, int(pick_max(2.0, 9.0))\nprint *, pick_max(.false., .true.)\nend program t\n",
        ["9", "9", "true"]
    };

    generic_len_overload_int_char => {
        "module glen\nimplicit none\ninterface span\nmodule procedure span_int, span_char\nend interface\ncontains\nfunction span_int(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = abs(n) + 1\nend function span_int\nfunction span_char(s) result(r)\ncharacter(len=*), intent(in) :: s\ninteger :: r\nr = len_trim(s)\nend function span_char\nend module glen\nprogram t\nuse glen\nprint *, span(-4)\nprint *, span('abcd')\nend program t\n",
        ["5", "4"]
    };

    generic_abs_int_real => {
        "module gabs\nimplicit none\ninterface my_abs\nmodule procedure abs_int, abs_real\nend interface\ncontains\nfunction abs_int(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nif (x < 0) then\nr = -x\nelse\nr = x\nend if\nend function abs_int\nfunction abs_real(x) result(r)\nreal, intent(in) :: x\nreal :: r\nr = abs(x)\nend function abs_real\nend module gabs\nprogram t\nuse gabs\nprint *, my_abs(-7)\nprint *, int(my_abs(-7.0))\nend program t\n",
        ["7", "7"]
    };

    generic_compare_strings_and_ints => {
        "module gcmp\nimplicit none\ninterface same\nmodule procedure same_int, same_char\nend interface\ncontains\nfunction same_int(a, b) result(r)\ninteger, intent(in) :: a, b\nlogical :: r\nr = a == b\nend function same_int\nfunction same_char(a, b) result(r)\ncharacter(len=*), intent(in) :: a, b\nlogical :: r\nr = a == b\nend function same_char\nend module gcmp\nprogram t\nuse gcmp\nprint *, same(3, 3)\nprint *, same('x', 'x')\nend program t\n",
        ["true", "true"]
    };

    generic_sum_real_array_and_pair => {
        "module gsum\nimplicit none\ninterface total\nmodule procedure total_pair, total_real2\nend interface\ncontains\nfunction total_pair(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = a + b\nend function total_pair\nfunction total_real2(x, y) result(r)\nreal, intent(in) :: x, y\nreal :: r\nr = x + y\nend function total_real2\nend module gsum\nprogram t\nuse gsum\nprint *, total(4, 5)\nprint *, int(total(1.5, 2.5))\nend program t\n",
        ["9", "4"]
    };

    // ── Module procedure interface resolution ────────────────────────

    module_interface_single_procedure => {
        "module giface\nimplicit none\ninterface clamp_int\nmodule procedure clamp_value\nend interface\ncontains\nfunction clamp_value(v, lo, hi) result(r)\ninteger, intent(in) :: v, lo, hi\ninteger :: r\nif (v < lo) then\nr = lo\nelse if (v > hi) then\nr = hi\nelse\nr = v\nend if\nend function clamp_value\nend module giface\nprogram t\nuse giface\nprint *, clamp_int(15, 0, 10)\nend program t\n",
        ["10"]
    };

    module_interface_two_procedures => {
        "module gmid\nimplicit none\ninterface middle\nmodule procedure mid_int, mid_real\nend interface\ncontains\nfunction mid_int(a, b) result(r)\ninteger, intent(in) :: a, b\ninteger :: r\nr = (a + b) / 2\nend function mid_int\nfunction mid_real(a, b) result(r)\nreal, intent(in) :: a, b\nreal :: r\nr = (a + b) / 2.0\nend function mid_real\nend module gmid\nprogram t\nuse gmid\nprint *, middle(3, 7)\nprint *, int(middle(3.0, 7.0))\nend program t\n",
        ["5", "5"]
    };

    module_interface_external_shape => {
        "module gext\nimplicit none\ninterface\nfunction extern_triple(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nend function extern_triple\nend interface\ncontains\nfunction via_iface(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = extern_triple(n)\nend function via_iface\nend module gext\nfunction extern_triple(n) result(r)\ninteger, intent(in) :: n\ninteger :: r\nr = n * 3\nend function extern_triple\nprogram t\nuse gext\nprint *, via_iface(4)\nend program t\n",
        ["12"]
    };

    module_interface_operator_plus_int_wrap => {
        "module giwrap\nimplicit none\ntype :: Wrap\ninteger :: v\nend type Wrap\ninterface operator(+)\nmodule procedure add_wrap\nend interface\ncontains\nfunction add_wrap(a, b) result(c)\ntype(Wrap), intent(in) :: a, b\ntype(Wrap) :: c\nc%v = a%v + b%v\nend function add_wrap\nend module giwrap\nprogram t\nuse giwrap\ntype(Wrap) :: a, b, c\na%v = 100\nb%v = 23\nc = a + b\nprint *, c%v\nend program t\n",
        ["123"]
    };

    module_interface_assignment_chain => {
        "module gchain\nimplicit none\ntype :: Cell\ninteger :: v\nend type Cell\ninterface assignment(=)\nmodule procedure set_cell\nend interface\ncontains\nsubroutine set_cell(dest, src)\ntype(Cell), intent(out) :: dest\ninteger, intent(in) :: src\ndest%v = src * 10\nend subroutine set_cell\nend module gchain\nprogram t\nuse gchain\ntype(Cell) :: c\nc = 6\nprint *, c%v\nend program t\n",
        ["60"]
    };

    // ── Combined operator and generic patterns ───────────────────────

    generic_operator_apply_twice => {
        "module gtwice\nimplicit none\ntype :: Num\ninteger :: v\nend type Num\ninterface operator(+)\nmodule procedure add_num\nend interface\ncontains\nfunction add_num(a, b) result(c)\ntype(Num), intent(in) :: a, b\ntype(Num) :: c\nc%v = a%v + b%v\nend function add_num\nend module gtwice\nprogram t\nuse gtwice\ntype(Num) :: a, b, c, d\na%v = 1\nb%v = 2\nc%v = 3\nd = a + b + c\nprint *, d%v\nend program t\n",
        ["6"]
    };

    assignment_then_operator_on_type => {
        "module gmix\nimplicit none\ntype :: Acc\ninteger :: v\nend type Acc\ninterface assignment(=)\nmodule procedure set_acc\nend interface\ninterface operator(+)\nmodule procedure add_acc\nend interface\ncontains\nsubroutine set_acc(dest, src)\ntype(Acc), intent(out) :: dest\ninteger, intent(in) :: src\ndest%v = src\nend subroutine set_acc\nfunction add_acc(a, b) result(c)\ntype(Acc), intent(in) :: a, b\ntype(Acc) :: c\nc%v = a%v + b%v\nend function add_acc\nend module gmix\nprogram t\nuse gmix\ntype(Acc) :: x, y, z\nx = 4\ny = 5\nz = x + y\nprint *, z%v\nend program t\n",
        ["9"]
    };

    generic_interface_three_int_procedures => {
        "module gthree\nimplicit none\ninterface pick\nmodule procedure pick_a, pick_b, pick_c\nend interface\ncontains\nfunction pick_a(v) result(r)\ninteger, intent(in) :: v\ninteger :: r\nr = v\nend function pick_a\nfunction pick_b(v) result(r)\ninteger, intent(in) :: v\ninteger :: r\nr = v + 1\nend function pick_b\nfunction pick_c(v) result(r)\ninteger, intent(in) :: v\ninteger :: r\nr = v + 2\nend function pick_c\nend module gthree\nprogram t\nuse gthree\nprint *, pick(1)\nend program t\n",
        ["1"]
    };

    operator_multiply_accumulate_boxes => {
        "module gacc\nimplicit none\ntype :: Box\ninteger :: v\nend type Box\ninterface operator(*)\nmodule procedure mul_box\nend interface\ncontains\nfunction mul_box(a, b) result(c)\ntype(Box), intent(in) :: a, b\ntype(Box) :: c\nc%v = a%v * b%v\nend function mul_box\nend module gacc\nprogram t\nuse gacc\ntype(Box) :: a, b, c, d\na%v = 2\nb%v = 3\nc%v = 4\nd = a * b * c\nprint *, d%v\nend program t\n",
        ["24"]
    };

    module_generic_logical_negation => {
        "module gnot\nimplicit none\ninterface flip\nmodule procedure flip_log, flip_int\nend interface\ncontains\nfunction flip_log(v) result(r)\nlogical, intent(in) :: v\nlogical :: r\nr = .not. v\nend function flip_log\nfunction flip_int(v) result(r)\ninteger, intent(in) :: v\ninteger :: r\nr = -v\nend function flip_int\nend module gnot\nprogram t\nuse gnot\nprint *, flip(.true.)\nprint *, flip(8)\nend program t\n",
        ["false", "-8"]
    };

    operator_concat_three_tokens => {
        "module g3tok\nimplicit none\ntype :: Tok\ncharacter(len=4) :: s\nend type Tok\ninterface operator(//)\nmodule procedure cat_tok\nend interface\ncontains\nfunction cat_tok(a, b) result(c)\ntype(Tok), intent(in) :: a, b\ntype(Tok) :: c\nc%s = trim(a%s) // trim(b%s)\nend function cat_tok\nend module g3tok\nprogram t\nuse g3tok\ntype(Tok) :: a, b, c, d\na%s = 'a'\nb%s = 'b'\nc%s = 'c'\nd = (a // b) // c\nprint *, trim(d%s)\nend program t\n",
        ["abc"]
    };

    generic_mixed_use_from_program => {
        "module gutil\nimplicit none\ninterface emit\nmodule procedure emit_int, emit_char\nend interface\ncontains\nsubroutine emit_int(v)\ninteger, intent(in) :: v\nprint *, v\nend subroutine emit_int\nsubroutine emit_char(s)\ncharacter(len=*), intent(in) :: s\nprint *, len_trim(s)\nend subroutine emit_char\nend module gutil\nprogram t\nuse gutil\ncall emit(9)\ncall emit('abcd')\nend program t\n",
        ["9", "4"]
    };

    module_interface_subroutine_generic => {
        "module gsub\nimplicit none\ninterface run\nmodule procedure run_int, run_real\nend interface\ncontains\nsubroutine run_int(v)\ninteger, intent(in) :: v\nprint *, v * 2\nend subroutine run_int\nsubroutine run_real(v)\nreal, intent(in) :: v\nprint *, int(v * 2.0)\nend subroutine run_real\nend module gsub\nprogram t\nuse gsub\ncall run(5)\ncall run(2.5)\nend program t\n",
        ["10", "5"]
    };

    operator_eq_on_character_wrapper => {
        "module gstr\nimplicit none\ntype :: Str\ncharacter(len=6) :: data\nend type Str\ninterface operator(==)\nmodule procedure eq_str\nend interface\ncontains\nfunction eq_str(a, b) result(r)\ntype(Str), intent(in) :: a, b\nlogical :: r\nr = trim(a%data) == trim(b%data)\nend function eq_str\nend module gstr\nprogram t\nuse gstr\ntype(Str) :: a, b\na%data = 'hi'\nb%data = 'hi'\nprint *, a == b\nend program t\n",
        ["true"]
    };

    assignment_real_pair_to_point => {
        "module gpt\nimplicit none\ntype :: Point\nreal :: x, y\nend type Point\ninterface assignment(=)\nmodule procedure pair_to_point\nend interface\ncontains\nsubroutine pair_to_point(dest, src)\ntype(Point), intent(out) :: dest\nreal, intent(in) :: src(2)\ndest%x = src(1)\ndest%y = src(2)\nend subroutine pair_to_point\nend module gpt\nprogram t\nuse gpt\ntype(Point) :: p\nreal :: vals(2)\nvals = [3.0, 4.0]\np = vals\nprint *, int(p%x)\nprint *, int(p%y)\nend program t\n",
        ["3", "4"]
    };

    generic_square_int_and_real => {
        "module gsqr\nimplicit none\ninterface square\nmodule procedure square_int, square_real\nend interface\ncontains\nfunction square_int(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * x\nend function square_int\nfunction square_real(x) result(r)\nreal, intent(in) :: x\nreal :: r\nr = x * x\nend function square_real\nend module gsqr\nprogram t\nuse gsqr\nprint *, square(6)\nprint *, int(square(2.5))\nend program t\n",
        ["36", "6"]
    };

    module_procedure_interface_in_nested_call => {
        "module gnest\nimplicit none\ninterface twice\nmodule procedure twice_val\nend interface\ncontains\nfunction twice_val(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = x * 2\nend function twice_val\nfunction quad(x) result(r)\ninteger, intent(in) :: x\ninteger :: r\nr = twice(x) + twice(x)\nend function quad\nend module gnest\nprogram t\nuse gnest\nprint *, quad(5)\nend program t\n",
        ["20"]
    };

    operator_minus_unary_on_metric => {
        "module gmet\nimplicit none\ntype :: Metric\ninteger :: mm\nend type Metric\ninterface operator(-)\nmodule procedure neg_metric\nend interface\ncontains\nfunction neg_metric(a) result(b)\ntype(Metric), intent(in) :: a\ntype(Metric) :: b\nb%mm = -a%mm\nend function neg_metric\nend module gmet\nprogram t\nuse gmet\ntype(Metric) :: m, n\nm%mm = 15\nn = -m\nprint *, n%mm\nend program t\n",
        ["-15"]
    };

    generic_identity_overloads => {
        "module gid\nimplicit none\ninterface ident\nmodule procedure ident_int, ident_char\nend interface\ncontains\nfunction ident_int(v) result(r)\ninteger, intent(in) :: v\ninteger :: r\nr = v\nend function ident_int\nfunction ident_char(v) result(r)\ncharacter(len=*), intent(in) :: v\ncharacter(len=10) :: r\nr = v\nend function ident_char\nend module gid\nprogram t\nuse gid\nprint *, ident(42)\nprint *, trim(ident('z'))\nend program t\n",
        ["42", "z"]
    };

    assignment_box_from_two_integers => {
        "module gpair\nimplicit none\ntype :: PairBox\ninteger :: a, b\nend type PairBox\ninterface assignment(=)\nmodule procedure ints_to_pairbox\nend interface\ncontains\nsubroutine ints_to_pairbox(dest, src)\ntype(PairBox), intent(out) :: dest\ninteger, intent(in) :: src(2)\ndest%a = src(1)\ndest%b = src(2)\nend subroutine ints_to_pairbox\nend module gpair\nprogram t\nuse gpair\ntype(PairBox) :: p\ninteger :: v(2)\nv = [2, 3]\np = v\nprint *, p%a + p%b\nend program t\n",
        ["5"]
    };

    module_interface_operator_on_accumulator => {
        "module gacc2\nimplicit none\ntype :: Acc\ninteger :: total\nend type Acc\ninterface operator(+)\nmodule procedure add_acc\nend interface\ncontains\nfunction add_acc(a, b) result(c)\ntype(Acc), intent(in) :: a, b\ntype(Acc) :: c\nc%total = a%total + b%total\nend function add_acc\nend module gacc2\nprogram t\nuse gacc2\ntype(Acc) :: seed, step, out\nseed%total = 10\nstep%total = 5\nout = seed + step\nprint *, out%total\nend program t\n",
        ["15"]
    };
}
