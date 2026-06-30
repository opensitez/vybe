//! Extended INTENT and OPTIONAL coverage: mixed kinds, multi-optional args,
//! PRESENT()-driven branches, default-when-absent patterns, and module wrappers.
//! Distinct from `test_subroutine_extended.rs` basic INTENT/OPTIONAL cases.

fortran_cases! {
    // ── INTENT(in): non-integer kinds ────────────────────────────────

    intent_in_real_hypotenuse_pair => {
        "program t\nprint *, rhyp(3.0, 4.0)\ncontains\nreal function rhyp(a, b)\nreal, intent(in) :: a, b\nrhyp = sqrt(a*a + b*b)\nend function rhyp\nend program t\n",
        ["5"]
    };

    intent_in_logical_conjunction => {
        "program t\nprint *, land(.true., .false.)\nprint *, land(.true., .true.)\ncontains\nlogical function land(p, q)\nlogical, intent(in) :: p, q\nland = p .and. q\nend function land\nend program t\n",
        ["false", "true"]
    };

    intent_in_character_concat_trim => {
        "program t\nprint *, trim(merge_names('Ada', 'Lovelace'))\ncontains\ncharacter(len=20) function merge_names(first, last)\ncharacter(len=*), intent(in) :: first, last\nmerge_names = trim(first) // ' ' // trim(last)\nend function merge_names\nend program t\n",
        ["Ada Lovelace"]
    };

    intent_in_complex_magnitude_squared => {
        "program t\nprint *, int(cmag2(3.0, 4.0))\ncontains\nreal function cmag2(re, im)\nreal, intent(in) :: re, im\ncmag2 = re*re + im*im\nend function cmag2\nend program t\n",
        ["25"]
    };

    intent_in_rank_one_stride_sum => {
        "program t\ninteger :: v(5)\nv = [1, 2, 3, 4, 5]\nprint *, stride_sum(v, 5, 2)\ncontains\nfunction stride_sum(a, n, step) result(s)\ninteger, intent(in) :: a(n), n, step\ninteger :: s, i\ns = 0\ndo i = 1, n, step\ns = s + a(i)\nend do\nend function stride_sum\nend program t\n",
        ["9"]
    };

    intent_in_parameter_passed_to_len => {
        "program t\nprint *, boxed_len('Fortran')\ncontains\ninteger function boxed_len(text)\ncharacter(len=*), intent(in) :: text\nboxed_len = len_trim(text) + 2\nend function boxed_len\nend program t\n",
        ["9"]
    };

    intent_in_double_precision_diff => {
        "program t\nprint *, int(dsub(5.5d0, 2.0d0))\ncontains\ndouble precision function dsub(a, b)\ndouble precision, intent(in) :: a, b\ndsub = a - b\nend function dsub\nend program t\n",
        ["3"]
    };

    intent_in_logical_equality_function => {
        "program t\nprint *, same_flag(.true., .true.)\nprint *, same_flag(.true., .false.)\ncontains\nlogical function same_flag(a, b)\nlogical, intent(in) :: a, b\nsame_flag = a .eqv. b\nend function same_flag\nend program t\n",
        ["true", "false"]
    };

    // ── INTENT(out): scalar and array outputs ────────────────────────

    intent_out_real_quarter_pi => {
        "program t\nreal :: x\ncall unit_quarter(x)\nprint *, x\ncontains\nsubroutine unit_quarter(v)\nreal, intent(out) :: v\nv = 0.25\nend subroutine unit_quarter\nend program t\n",
        ["0.25"]
    };

    intent_out_logical_success_flag => {
        "program t\nlogical :: ok\ncall set_ok(ok)\nprint *, ok\ncontains\nsubroutine set_ok(flag)\nlogical, intent(out) :: flag\nflag = .true.\nend subroutine set_ok\nend program t\n",
        ["true"]
    };

    intent_out_character_buffer_fill => {
        "program t\ncharacter(len=5) :: tag\ncall fill_tag(tag)\nprint *, trim(tag)\ncontains\nsubroutine fill_tag(s)\ncharacter(len=5), intent(out) :: s\ns = 'hello'\nend subroutine fill_tag\nend program t\n",
        ["hello"]
    };

    intent_out_three_integer_slots => {
        "program t\ninteger :: a, b, c\ncall triple_out(a, b, c)\nprint *, a + b + c\ncontains\nsubroutine triple_out(x, y, z)\ninteger, intent(out) :: x, y, z\nx = 2\ny = 3\nz = 4\nend subroutine triple_out\nend program t\n",
        ["9"]
    };

    intent_out_rank_two_shape_init => {
        "program t\ninteger :: m(2, 2)\ncall identity2(m)\nprint *, sum(m)\ncontains\nsubroutine identity2(a)\ninteger, intent(out) :: a(2, 2)\na = 0\na(1, 1) = 1\na(2, 2) = 1\nend subroutine identity2\nend program t\n",
        ["2"]
    };

    intent_out_complex_components => {
        "program t\nreal :: re, im\ncall unit_imag(re, im)\nprint *, int(re)\nprint *, int(im)\ncontains\nsubroutine unit_imag(r, i)\nreal, intent(out) :: r, i\nr = 0.0\ni = 1.0\nend subroutine unit_imag\nend program t\n",
        ["0", "1"]
    };

    intent_out_pair_from_quotient => {
        "program t\ninteger :: q, r\ncall divide_out(17, 5, q, r)\nprint *, q\nprint *, r\ncontains\nsubroutine divide_out(n, d, quot, rem)\ninteger, intent(in) :: n, d\ninteger, intent(out) :: quot, rem\nquot = n / d\nrem = mod(n, d)\nend subroutine divide_out\nend program t\n",
        ["3", "2"]
    };

    intent_out_substring_target => {
        "program t\ncharacter(len=4) :: code\ncall code_out(code)\nprint *, trim(code)\ncontains\nsubroutine code_out(c)\ncharacter(len=4), intent(out) :: c\nc = 'ABCD'\nend subroutine code_out\nend program t\n",
        ["ABCD"]
    };

    // ── INTENT(inout): mutate caller storage ─────────────────────────

    intent_inout_real_halve_value => {
        "program t\nreal :: x\nx = 8.0\ncall halve_real(x)\nprint *, x\ncontains\nsubroutine halve_real(v)\nreal, intent(inout) :: v\nv = v / 2.0\nend subroutine halve_real\nend program t\n",
        ["4"]
    };

    intent_inout_logical_toggle => {
        "program t\nlogical :: flag\nflag = .false.\ncall flip(flag)\nprint *, flag\ncontains\nsubroutine flip(v)\nlogical, intent(inout) :: v\nv = .not. v\nend subroutine flip\nend program t\n",
        ["true"]
    };

    intent_inout_character_first_char => {
        "program t\ncharacter(len=4) :: word\nword = 'test'\ncall upper_first(word)\nprint *, trim(word)\ncontains\nsubroutine upper_first(s)\ncharacter(len=4), intent(inout) :: s\nif (s(1:1) == 't') s(1:1) = 'T'\nend subroutine upper_first\nend program t\n",
        ["Test"]
    };

    intent_inout_increment_by_delta => {
        "program t\ninteger :: n\nn = 10\ncall bump_by(n, 7)\nprint *, n\ncontains\nsubroutine bump_by(x, delta)\ninteger, intent(inout) :: x\ninteger, intent(in) :: delta\nx = x + delta\nend subroutine bump_by\nend program t\n",
        ["17"]
    };

    intent_inout_reverse_two_element => {
        "program t\ninteger :: pair(2)\npair = [3, 9]\ncall reverse2(pair)\nprint *, pair(1)\nprint *, pair(2)\ncontains\nsubroutine reverse2(v)\ninteger, intent(inout) :: v(2)\ninteger :: t\nt = v(1)\nv(1) = v(2)\nv(2) = t\nend subroutine reverse2\nend program t\n",
        ["9", "3"]
    };

    intent_inout_scale_array_in_place => {
        "program t\ninteger :: a(3)\na = [2, 4, 6]\ncall scale_inplace(a, 3)\nprint *, sum(a)\ncontains\nsubroutine scale_inplace(v, n)\ninteger, intent(inout) :: v(n)\ninteger, intent(in) :: n\ninteger :: i\ndo i = 1, n\nv(i) = v(i) * 2\nend do\nend subroutine scale_inplace\nend program t\n",
        ["24"]
    };

    intent_inout_clamp_to_range => {
        "program t\ninteger :: v\nv = 15\ncall clamp(v, 0, 10)\nprint *, v\ncontains\nsubroutine clamp(x, lo, hi)\ninteger, intent(inout) :: x\ninteger, intent(in) :: lo, hi\nif (x < lo) x = lo\nif (x > hi) x = hi\nend subroutine clamp\nend program t\n",
        ["10"]
    };

    intent_inout_double_precision_add => {
        "program t\ndouble precision :: x\nx = 1.5d0\ncall add_half(x)\nprint *, int(x)\ncontains\nsubroutine add_half(v)\ndouble precision, intent(inout) :: v\nv = v + 0.5d0\nend subroutine add_half\nend program t\n",
        ["2"]
    };

    // ── OPTIONAL: single absent/present branches ─────────────────────

    optional_real_tolerance_default_zero => {
        "program t\nprint *, near(5.0, 5.0)\nprint *, near(5.0, 4.0, 0.5)\ncontains\nlogical function near(a, b, tol)\nreal, intent(in) :: a, b\nreal, intent(in), optional :: tol\nreal :: use_tol\nif (present(tol)) then\nuse_tol = tol\nelse\nuse_tol = 0.0\nend if\nnear = abs(a - b) <= use_tol\nend function near\nend program t\n",
        ["true", "false"]
    };

    optional_logical_invert_when_present => {
        "program t\nprint *, maybe_not(.true.)\nprint *, maybe_not(.true., .true.)\ncontains\nlogical function maybe_not(v, flip)\nlogical, intent(in) :: v\nlogical, intent(in), optional :: flip\nif (present(flip)) then\nmaybe_not = .not. v\nelse\nmaybe_not = v\nend if\nend function maybe_not\nend program t\n",
        ["true", "false"]
    };

    optional_character_suffix_missing => {
        "program t\nprint *, trim(label('core'))\ncontains\ncharacter(len=20) function label(base, suffix)\ncharacter(len=*), intent(in) :: base\ncharacter(len=*), intent(in), optional :: suffix\nif (present(suffix)) then\nlabel = trim(base) // '_' // trim(suffix)\nelse\nlabel = trim(base)\nend if\nend function label\nend program t\n",
        ["core"]
    };

    optional_character_suffix_present => {
        "program t\nprint *, trim(label('core', 'ext'))\ncontains\ncharacter(len=20) function label(base, suffix)\ncharacter(len=*), intent(in) :: base\ncharacter(len=*), intent(in), optional :: suffix\nif (present(suffix)) then\nlabel = trim(base) // '_' // trim(suffix)\nelse\nlabel = trim(base)\nend if\nend function label\nend program t\n",
        ["core_ext"]
    };

    optional_offset_defaults_to_zero => {
        "program t\nprint *, shift_val(8)\nprint *, shift_val(8, 3)\ncontains\ninteger function shift_val(x, off)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: off\nif (present(off)) then\nshift_val = x + off\nelse\nshift_val = x\nend if\nend function shift_val\nend program t\n",
        ["8", "11"]
    };

    optional_divisor_defaults_to_one => {
        "program t\nprint *, safe_div(12)\nprint *, safe_div(12, 4)\ncontains\ninteger function safe_div(n, d)\ninteger, intent(in) :: n\ninteger, intent(in), optional :: d\ninteger :: use_d\nif (present(d)) then\nuse_d = d\nelse\nuse_d = 1\nend if\nsafe_div = n / use_d\nend function safe_div\nend program t\n",
        ["12", "3"]
    };

    optional_subroutine_padding_absent => {
        "program t\ncall emit_padded(7)\ncontains\nsubroutine emit_padded(v, pad)\ninteger, intent(in) :: v\ninteger, intent(in), optional :: pad\ninteger :: out\nout = v\nif (present(pad)) out = out + pad\nprint *, out\nend subroutine emit_padded\nend program t\n",
        ["7"]
    };

    optional_subroutine_padding_present => {
        "program t\ncall emit_padded(7, 5)\ncontains\nsubroutine emit_padded(v, pad)\ninteger, intent(in) :: v\ninteger, intent(in), optional :: pad\ninteger :: out\nout = v\nif (present(pad)) out = out + pad\nprint *, out\nend subroutine emit_padded\nend program t\n",
        ["12"]
    };

    optional_inout_addend_only_when_present => {
        "program t\ninteger :: n\nn = 4\ncall maybe_add(n)\nprint *, n\ncontains\nsubroutine maybe_add(x, extra)\ninteger, intent(inout) :: x\ninteger, intent(in), optional :: extra\nif (present(extra)) x = x + extra\nend subroutine maybe_add\nend program t\n",
        ["4"]
    };

    optional_inout_addend_applied => {
        "program t\ninteger :: n\nn = 4\ncall maybe_add(n, 6)\nprint *, n\ncontains\nsubroutine maybe_add(x, extra)\ninteger, intent(inout) :: x\ninteger, intent(in), optional :: extra\nif (present(extra)) x = x + extra\nend subroutine maybe_add\nend program t\n",
        ["10"]
    };

    // ── OPTIONAL: multiple dummy arguments ───────────────────────────

    optional_two_both_absent => {
        "program t\nprint *, combine3(4)\ncontains\ninteger function combine3(a, b, c)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b, c\ncombine3 = a\nif (present(b)) combine3 = combine3 + b\nif (present(c)) combine3 = combine3 + c\nend function combine3\nend program t\n",
        ["4"]
    };

    optional_two_first_present => {
        "program t\nprint *, combine3(4, 2)\ncontains\ninteger function combine3(a, b, c)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b, c\ncombine3 = a\nif (present(b)) combine3 = combine3 + b\nif (present(c)) combine3 = combine3 + c\nend function combine3\nend program t\n",
        ["6"]
    };

    optional_two_both_present => {
        "program t\nprint *, combine3(4, 2, 1)\ncontains\ninteger function combine3(a, b, c)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b, c\ncombine3 = a\nif (present(b)) combine3 = combine3 + b\nif (present(c)) combine3 = combine3 + c\nend function combine3\nend program t\n",
        ["7"]
    };

    optional_middle_arg_present_only => {
        "program t\nprint *, tri_opt(1, 5)\ncontains\ninteger function tri_opt(a, b, c)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b, c\ninteger :: s\ns = a\nif (present(b)) s = s + b\nif (present(c)) s = s + c\ntri_opt = s\nend function tri_opt\nend program t\n",
        ["6"]
    };

    optional_three_tail_args_partial => {
        "program t\nprint *, quad_sum(1, 2)\ncontains\ninteger function quad_sum(w, x, y, z)\ninteger, intent(in) :: w\ninteger, intent(in), optional :: x, y, z\nquad_sum = w\nif (present(x)) quad_sum = quad_sum + x\nif (present(y)) quad_sum = quad_sum + y\nif (present(z)) quad_sum = quad_sum + z\nend function quad_sum\nend program t\n",
        ["3"]
    };

    optional_three_tail_args_all => {
        "program t\nprint *, quad_sum(1, 2, 3, 4)\ncontains\ninteger function quad_sum(w, x, y, z)\ninteger, intent(in) :: w\ninteger, intent(in), optional :: x, y, z\nquad_sum = w\nif (present(x)) quad_sum = quad_sum + x\nif (present(y)) quad_sum = quad_sum + y\nif (present(z)) quad_sum = quad_sum + z\nend function quad_sum\nend program t\n",
        ["10"]
    };

    optional_pair_subroutine_flags => {
        "program t\ncall report_flags(3)\ncall report_flags(3, 2)\ncall report_flags(3, 2, 1)\ncontains\nsubroutine report_flags(a, b, c)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b, c\ninteger :: s\ns = a\nif (present(b)) s = s + b\nif (present(c)) s = s + c\nprint *, s\nend subroutine report_flags\nend program t\n",
        ["3", "5", "6"]
    };

    optional_last_of_four_missing => {
        "program t\nprint *, pack4(1, 2, 3)\ncontains\ninteger function pack4(a, b, c, d)\ninteger, intent(in) :: a, b, c\ninteger, intent(in), optional :: d\npack4 = a + b + c\nif (present(d)) pack4 = pack4 + d\nend function pack4\nend program t\n",
        ["6"]
    };

    optional_last_of_four_present => {
        "program t\nprint *, pack4(1, 2, 3, 4)\ncontains\ninteger function pack4(a, b, c, d)\ninteger, intent(in) :: a, b, c\ninteger, intent(in), optional :: d\npack4 = a + b + c\nif (present(d)) pack4 = pack4 + d\nend function pack4\nend program t\n",
        ["10"]
    };

    optional_real_pair_weighted => {
        "program t\nprint *, int(wavg(2.0, 8.0))\nprint *, int(wavg(2.0, 8.0, 0.25))\ncontains\nreal function wavg(a, b, w)\nreal, intent(in) :: a, b\nreal, intent(in), optional :: w\nreal :: use_w\nif (present(w)) then\nuse_w = w\nelse\nuse_w = 0.5\nend if\nwavg = a * use_w + b * (1.0 - use_w)\nend function wavg\nend program t\n",
        ["5", "6"]
    };

    // ── PRESENT(): explicit branch reporting ─────────────────────────

    present_flag_reports_zero_when_absent => {
        "program t\ncall show_present(4)\ncontains\nsubroutine show_present(x, y)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: y\nif (present(y)) then\nprint *, 1\nelse\nprint *, 0\nend if\nend subroutine show_present\nend program t\n",
        ["0"]
    };

    present_flag_reports_one_when_present => {
        "program t\ncall show_present(4, 9)\ncontains\nsubroutine show_present(x, y)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: y\nif (present(y)) then\nprint *, 1\nelse\nprint *, 0\nend if\nend subroutine show_present\nend program t\n",
        ["1"]
    };

    present_drives_merge_style_default => {
        "program t\nprint *, pick_or(7, 3)\nprint *, pick_or(7)\ncontains\ninteger function pick_or(base, alt)\ninteger, intent(in) :: base\ninteger, intent(in), optional :: alt\nif (present(alt)) then\npick_or = alt\nelse\npick_or = base\nend if\nend function pick_or\nend program t\n",
        ["3", "7"]
    };

    present_guard_skips_optional_read => {
        "program t\nprint *, guarded_add(5)\nprint *, guarded_add(5, 8)\ncontains\ninteger function guarded_add(a, b)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b\nguarded_add = a\nif (present(b)) guarded_add = guarded_add + b\nend function guarded_add\nend program t\n",
        ["5", "13"]
    };

    present_on_character_optional => {
        "program t\ncall tag_len('abc')\ncall tag_len('abc', 'xyz')\ncontains\nsubroutine tag_len(a, b)\ncharacter(len=*), intent(in) :: a\ncharacter(len=*), intent(in), optional :: b\nif (present(b)) then\nprint *, len_trim(a) + len_trim(b)\nelse\nprint *, len_trim(a)\nend if\nend subroutine tag_len\nend program t\n",
        ["3", "6"]
    };

    present_on_logical_optional => {
        "program t\nprint *, opt_and(.true.)\nprint *, opt_and(.true., .false.)\ncontains\nlogical function opt_and(a, b)\nlogical, intent(in) :: a\nlogical, intent(in), optional :: b\nif (present(b)) then\nopt_and = a .and. b\nelse\nopt_and = a\nend if\nend function opt_and\nend program t\n",
        ["true", "false"]
    };

    present_count_two_optional_args => {
        "program t\nprint *, count_present(1)\nprint *, count_present(1, 2)\nprint *, count_present(1, 2, 3)\ncontains\ninteger function count_present(a, b, c)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b, c\ncount_present = 0\nif (present(b)) count_present = count_present + 1\nif (present(c)) count_present = count_present + 1\nend function count_present\nend program t\n",
        ["0", "1", "2"]
    };

    present_with_intent_out_optional => {
        "program t\ninteger :: r\ncall maybe_fill(5, r)\nprint *, r\ncontains\nsubroutine maybe_fill(seed, outv, use_out)\ninteger, intent(in) :: seed\ninteger, intent(out) :: outv\nlogical, intent(in), optional :: use_out\nif (present(use_out) .and. use_out) then\noutv = seed * 2\nelse\noutv = seed\nend if\nend subroutine maybe_fill\nend program t\n",
        ["5"]
    };

    // ── Default-when-absent value patterns ───────────────────────────

    default_cap_when_max_absent => {
        "program t\nprint *, capped(12)\nprint *, capped(12, 10)\ncontains\ninteger function capped(v, lim)\ninteger, intent(in) :: v\ninteger, intent(in), optional :: lim\ninteger :: use_lim\nif (present(lim)) then\nuse_lim = lim\nelse\nuse_lim = 100\nend if\nif (v > use_lim) then\ncapped = use_lim\nelse\ncapped = v\nend if\nend function capped\nend program t\n",
        ["12", "10"]
    };

    default_floor_when_min_absent => {
        "program t\nprint *, floored(-3)\nprint *, floored(-3, 0)\ncontains\ninteger function floored(v, lo)\ninteger, intent(in) :: v\ninteger, intent(in), optional :: lo\ninteger :: use_lo\nif (present(lo)) then\nuse_lo = lo\nelse\nuse_lo = 0\nend if\nif (v < use_lo) then\nfloored = use_lo\nelse\nfloored = v\nend if\nend function floored\nend program t\n",
        ["0", "0"]
    };

    default_repeat_count_one => {
        "program t\nprint *, repeat_val(9)\nprint *, repeat_val(9, 3)\ncontains\ninteger function repeat_val(x, n)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: n\ninteger :: use_n, i\nif (present(n)) then\nuse_n = n\nelse\nuse_n = 1\nend if\nrepeat_val = 0\ndo i = 1, use_n\nrepeat_val = repeat_val + x\nend do\nend function repeat_val\nend program t\n",
        ["9", "27"]
    };

    default_string_placeholder => {
        "program t\nprint *, trim(name_or('Ada'))\nprint *, trim(name_or('Ada', 'Unknown'))\ncontains\ncharacter(len=20) function name_or(got, fallback)\ncharacter(len=*), intent(in) :: got\ncharacter(len=*), intent(in), optional :: fallback\nif (present(fallback)) then\nname_or = fallback\nelse\nname_or = got\nend if\nend function name_or\nend program t\n",
        ["Ada", "Unknown"]
    };

    default_scale_factor_ten => {
        "program t\nprint *, scaled(3)\nprint *, scaled(3, 5)\ncontains\ninteger function scaled(v, factor)\ninteger, intent(in) :: v\ninteger, intent(in), optional :: factor\ninteger :: use_f\nif (present(factor)) then\nuse_f = factor\nelse\nuse_f = 10\nend if\nscaled = v * use_f\nend function scaled\nend program t\n",
        ["30", "15"]
    };

    default_boolean_or_true => {
        "program t\nprint *, coalesce_bool(.false.)\nprint *, coalesce_bool(.false., .true.)\ncontains\nlogical function coalesce_bool(v, alt)\nlogical, intent(in) :: v\nlogical, intent(in), optional :: alt\nif (present(alt)) then\ncoalesce_bool = alt\nelse\ncoalesce_bool = .true.\nend if\nend function coalesce_bool\nend program t\n",
        ["true", "true"]
    };

    default_real_unit_value => {
        "program t\nprint *, with_unit(4.0)\nprint *, with_unit(4.0, 2.0)\ncontains\nreal function with_unit(v, u)\nreal, intent(in) :: v\nreal, intent(in), optional :: u\nreal :: use_u\nif (present(u)) then\nuse_u = u\nelse\nuse_u = 1.0\nend if\nwith_unit = v * use_u\nend function with_unit\nend program t\n",
        ["4", "8"]
    };

    default_power_exponent_two => {
        "program t\nprint *, opt_pow(5)\nprint *, opt_pow(5, 3)\ncontains\ninteger function opt_pow(base, exp)\ninteger, intent(in) :: base\ninteger, intent(in), optional :: exp\ninteger :: use_e, i\nif (present(exp)) then\nuse_e = exp\nelse\nuse_e = 2\nend if\nopt_pow = 1\ndo i = 1, use_e\nopt_pow = opt_pow * base\nend do\nend function opt_pow\nend program t\n",
        ["25", "125"]
    };

    // ── Module wrappers and optional arrays ────────────────────────────

    module_optional_wrapper_defaults => {
        "module optwrap\nimplicit none\ncontains\nfunction bump(x, inc) result(r)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: inc\ninteger :: r\nif (present(inc)) then\nr = x + inc\nelse\nr = x + 1\nend if\nend function bump\nend module optwrap\nprogram t\nuse optwrap\nprint *, bump(10)\nprint *, bump(10, 4)\nend program t\n",
        ["11", "14"]
    };

    module_intent_out_pair_via_call => {
        "module outp\nimplicit none\ncontains\nsubroutine split10(a, b)\ninteger, intent(out) :: a, b\na = 4\nb = 6\nend subroutine split10\nend module outp\nprogram t\nuse outp\ninteger :: x, y\ncall split10(x, y)\nprint *, x + y\nend program t\n",
        ["10"]
    };

    optional_vector_sum_missing => {
        "program t\ninteger :: v(3)\nv = [1, 2, 3]\nprint *, sum_opt(v, 3)\ncontains\ninteger function sum_opt(a, n, extra)\ninteger, intent(in) :: a(n), n\ninteger, intent(in), optional :: extra\ninteger :: i\nsum_opt = 0\ndo i = 1, n\nsum_opt = sum_opt + a(i)\nend do\nif (present(extra)) sum_opt = sum_opt + extra\nend function sum_opt\nend program t\n",
        ["6"]
    };

    optional_vector_sum_with_extra => {
        "program t\ninteger :: v(3)\nv = [1, 2, 3]\nprint *, sum_opt(v, 3, 10)\ncontains\ninteger function sum_opt(a, n, extra)\ninteger, intent(in) :: a(n), n\ninteger, intent(in), optional :: extra\ninteger :: i\nsum_opt = 0\ndo i = 1, n\nsum_opt = sum_opt + a(i)\nend do\nif (present(extra)) sum_opt = sum_opt + extra\nend function sum_opt\nend program t\n",
        ["16"]
    };

    optional_array_copy_when_present => {
        "program t\ninteger :: src(2), dst(2)\nsrc = [4, 5]\ncall maybe_copy(src, dst, 2)\nprint *, sum(dst)\ncontains\nsubroutine maybe_copy(from, to, n, enable)\ninteger, intent(in) :: from(n), n\ninteger, intent(inout) :: to(n)\nlogical, intent(in), optional :: enable\ninteger :: i\nif (present(enable) .and. enable) then\ndo i = 1, n\nto(i) = from(i)\nend do\nelse\nto = 0\nend if\nend subroutine maybe_copy\nend program t\n",
        ["0"]
    };

    optional_array_copy_enabled => {
        "program t\ninteger :: src(2), dst(2)\nsrc = [4, 5]\ncall maybe_copy(src, dst, 2, .true.)\nprint *, sum(dst)\ncontains\nsubroutine maybe_copy(from, to, n, enable)\ninteger, intent(in) :: from(n), n\ninteger, intent(inout) :: to(n)\nlogical, intent(in), optional :: enable\ninteger :: i\nif (present(enable) .and. enable) then\ndo i = 1, n\nto(i) = from(i)\nend do\nelse\nto = 0\nend if\nend subroutine maybe_copy\nend program t\n",
        ["9"]
    };

    intent_inout_optional_scale_array => {
        "program t\ninteger :: a(2)\na = [3, 4]\ncall scale_if(a, 2, 2)\nprint *, sum(a)\ncontains\nsubroutine scale_if(v, n, k)\ninteger, intent(inout) :: v(n)\ninteger, intent(in) :: n\ninteger, intent(in), optional :: k\ninteger :: i\nif (present(k)) then\ndo i = 1, n\nv(i) = v(i) * k\nend do\nend if\nend subroutine scale_if\nend program t\n",
        ["14"]
    };

    intent_inout_optional_scale_applied => {
        "program t\ninteger :: a(2)\na = [3, 4]\ncall scale_if(a, 2, 3)\nprint *, sum(a)\ncontains\nsubroutine scale_if(v, n, k)\ninteger, intent(inout) :: v(n)\ninteger, intent(in) :: n\ninteger, intent(in), optional :: k\ninteger :: i\nif (present(k)) then\ndo i = 1, n\nv(i) = v(i) * k\nend do\nend if\nend subroutine scale_if\nend program t\n",
        ["21"]
    };

    optional_chain_through_internal_wrapper => {
        "program t\nprint *, wrap_add(2)\nprint *, wrap_add(2, 8)\ncontains\nfunction inner_add(a, b) result(r)\ninteger, intent(in) :: a\ninteger, intent(in), optional :: b\ninteger :: r\nr = a\nif (present(b)) r = r + b\nend function inner_add\ninteger function wrap_add(x, y)\ninteger, intent(in) :: x\ninteger, intent(in), optional :: y\nif (present(y)) then\nwrap_add = inner_add(x, y)\nelse\nwrap_add = inner_add(x)\nend if\nend function wrap_add\nend program t\n",
        ["2", "10"]
    };
}
