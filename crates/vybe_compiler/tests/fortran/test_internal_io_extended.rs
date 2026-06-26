//! Extended Fortran internal I/O: read/write to character variables, formatted
//! internal read, list-directed internal, and parsing integers/reals from strings.
//! Distinct from `test_internal_io.rs` (basic write/read roundtrips and helpers).

fortran_cases! {
    // ── List-directed WRITE to character buffer ──────────────────────

    iio_write_negative_int_star => {
        "program t\ncharacter(len=12) :: buf\nwrite(buf, *) -99\nprint *, trim(adjustl(buf))\nend program t\n",
        ["-99"]
    };

    iio_write_zero_i0_descriptor => {
        "program t\ncharacter(len=4) :: buf\nwrite(buf, '(I0)') 0\nprint *, trim(buf)\nend program t\n",
        ["0"]
    };

    iio_write_large_integer_i0 => {
        "program t\ncharacter(len=12) :: buf\nwrite(buf, '(I0)') 987654321\nprint *, trim(buf)\nend program t\n",
        ["987654321"]
    };

    iio_write_real_literal_star => {
        "program t\ncharacter(len=16) :: buf\nwrite(buf, *) 2.5\nprint *, trim(adjustl(buf))\nend program t\n",
        ["2.5"]
    };

    iio_write_string_literal_star => {
        "program t\ncharacter(len=12) :: buf\nwrite(buf, *) 'data'\nprint *, trim(adjustl(buf))\nend program t\n",
        ["data"]
    };

    iio_write_logical_true_star => {
        "program t\ncharacter(len=8) :: buf\nwrite(buf, *) .true.\nprint *, trim(adjustl(buf))\nend program t\n",
        ["true"]
    };

    iio_write_logical_false_star => {
        "program t\ncharacter(len=8) :: buf\nwrite(buf, *) .false.\nprint *, trim(adjustl(buf))\nend program t\n",
        ["false"]
    };

    iio_write_two_integers_star => {
        "program t\ncharacter(len=16) :: buf\nwrite(buf, *) 8, 9\nprint *, index(buf, '8')\nend program t\n",
        ["1"]
    };

    iio_write_named_unit_and_fmt => {
        "program t\ncharacter(len=6) :: buf\nwrite(unit=buf, fmt='(I0)') 13\nprint *, trim(buf)\nend program t\n",
        ["13"]
    };

    iio_write_to_character_array_slot => {
        "program t\ncharacter(len=6) :: slots(3)\nwrite(slots(2), '(I0)') 24\nprint *, trim(slots(2))\nend program t\n",
        ["24"]
    };

    // ── Formatted WRITE to character buffer ──────────────────────────

    iio_write_formatted_i4_padded => {
        "program t\ncharacter(len=6) :: buf\nwrite(buf, '(I4)') 7\nprint *, buf(1:4)\nend program t\n",
        ["   7"]
    };

    iio_write_formatted_f62 => {
        "program t\ncharacter(len=8) :: buf\nwrite(buf, '(F6.2)') 1.25\nprint *, trim(buf)\nend program t\n",
        ["  1.25"]
    };

    iio_write_formatted_a_width => {
        "program t\ncharacter(len=10) :: buf\nwrite(buf, '(A8)') 'Fortran'\nprint *, buf(1:8)\nend program t\n",
        ["Fortran "]
    };

    iio_write_formatted_l1_false => {
        "program t\ncharacter(len=4) :: buf\nwrite(buf, '(L1)') .false.\nprint *, buf(1:1)\nend program t\n",
        ["F"]
    };

    iio_write_formatted_es_exponent => {
        "program t\ncharacter(len=14) :: buf\nwrite(buf, '(ES9.2)') 0.0045\nprint *, len_trim(buf)\nend program t\n",
        ["9"]
    };

    iio_write_formatted_b_binary => {
        "program t\ncharacter(len=8) :: buf\nwrite(buf, '(B8)') 15\nprint *, trim(buf)\nend program t\n",
        ["00001111"]
    };

    iio_write_formatted_o_octal => {
        "program t\ncharacter(len=4) :: buf\nwrite(buf, '(O4)') 10\nprint *, trim(buf)\nend program t\n",
        ["0012"]
    };

    iio_write_formatted_z_hex => {
        "program t\ncharacter(len=6) :: buf\nwrite(buf, '(Z4)') 255\nprint *, trim(buf)\nend program t\n",
        ["00FF"]
    };

    iio_write_mixed_i_and_a_format => {
        "program t\ncharacter(len=12) :: buf\nwrite(buf, '(I0,A,I0)') 3, '-', 7\nprint *, trim(buf)\nend program t\n",
        ["3-7"]
    };

    iio_write_repeated_i_descriptor => {
        "program t\ncharacter(len=10) :: buf\nwrite(buf, '(2I4)') 11, 22\nprint *, buf(1:8)\nend program t\n",
        ["  11  22"]
    };

    iio_write_advance_no_two_fields => {
        "program t\ncharacter(len=8) :: buf = '        '\nwrite(buf(1:4), '(I2)', advance='no') 4\nwrite(buf(3:6), '(I2)', advance='no') 5\nprint *, buf(1:4)\nend program t\n",
        [" 445"]
    };

    iio_write_double_precision_d => {
        "program t\ncharacter(len=20) :: buf\nreal(kind=8) :: x = 6.25d0\nwrite(buf, '(F6.2)') x\nprint *, trim(buf)\nend program t\n",
        ["  6.25"]
    };

    // ── List-directed READ from character buffer ─────────────────────

    iio_read_negative_integer_star => {
        "program t\ncharacter(len=8) :: buf = '-42'\ninteger :: n\nread(buf, *) n\nprint *, n\nend program t\n",
        ["-42"]
    };

    iio_read_leading_spaces_star => {
        "program t\ncharacter(len=10) :: buf = '    17'\ninteger :: n\nread(buf, *) n\nprint *, n\nend program t\n",
        ["17"]
    };

    iio_read_zero_value_star => {
        "program t\ncharacter(len=6) :: buf = '0'\ninteger :: n\nread(buf, *) n\nprint *, n\nend program t\n",
        ["0"]
    };

    iio_read_real_literal_star => {
        "program t\ncharacter(len=10) :: buf = '3.5'\nreal :: x\nread(buf, *) x\nprint *, x\nend program t\n",
        ["3.5"]
    };

    iio_read_negative_real_star => {
        "program t\ncharacter(len=10) :: buf = '-1.5'\nreal :: x\nread(buf, *) x\nprint *, x\nend program t\n",
        ["-1.5"]
    };

    iio_read_scientific_notation_star => {
        "program t\ncharacter(len=12) :: buf = '1.5e2'\nreal :: x\nread(buf, *) x\nprint *, int(x)\nend program t\n",
        ["150"]
    };

    iio_read_three_integers_sum => {
        "program t\ncharacter(len=16) :: buf = '4 5 6'\ninteger :: a, b, c\nread(buf, *) a, b, c\nprint *, a + b + c\nend program t\n",
        ["15"]
    };

    iio_read_integer_and_real => {
        "program t\ncharacter(len=12) :: buf = '3 2.5'\ninteger :: i\nreal :: x\nread(buf, *) i, x\nprint *, i + int(x)\nend program t\n",
        ["5"]
    };

    iio_read_logical_dot_true => {
        "program t\ncharacter(len=12) :: buf = '.true.'\nlogical :: flag\nread(buf, *) flag\nprint *, flag\nend program t\n",
        ["true"]
    };

    iio_read_logical_dot_false => {
        "program t\ncharacter(len=12) :: buf = '.false.'\nlogical :: flag\nread(buf, *) flag\nprint *, flag\nend program t\n",
        ["false"]
    };

    // ── Formatted READ from character buffer ───────────────────────

    iio_read_formatted_i3_padded => {
        "program t\ncharacter(len=6) :: buf = '007'\ninteger :: n\nread(buf, '(I3)') n\nprint *, n\nend program t\n",
        ["7"]
    };

    iio_read_formatted_f52 => {
        "program t\ncharacter(len=10) :: buf = ' 2.50'\nreal :: x\nread(buf, '(F5.2)') x\nprint *, x\nend program t\n",
        ["2.5"]
    };

    iio_read_formatted_a_substring => {
        "program t\ncharacter(len=12) :: buf = 'alpha beta'\ncharacter(len=5) :: word\nread(buf, '(A5)') word\nprint *, word\nend program t\n",
        ["alpha"]
    };

    iio_read_formatted_two_i4_fields => {
        "program t\ncharacter(len=12) :: buf = '  3  14'\ninteger :: a, b\nread(buf, '(2I4)') a, b\nprint *, a + b\nend program t\n",
        ["17"]
    };

    iio_read_formatted_skip_x_descriptor => {
        "program t\ncharacter(len=10) :: buf = 'ab12'\ninteger :: n\nread(buf, '(2X,I2)') n\nprint *, n\nend program t\n",
        ["12"]
    };

    iio_read_formatted_comma_separated => {
        "program t\ncharacter(len=10) :: buf = '2,4'\ninteger :: a, b\nread(buf, '(I0,\",\",I0)') a, b\nprint *, a * b\nend program t\n",
        ["8"]
    };

    iio_read_formatted_l1_true => {
        "program t\ncharacter(len=4) :: buf = 'T   '\nlogical :: flag\nread(buf, '(L1)') flag\nprint *, flag\nend program t\n",
        ["true"]
    };

    iio_read_double_precision_d_suffix => {
        "program t\ncharacter(len=12) :: buf = '2.0d1'\nreal(kind=8) :: d\nread(buf, *) d\nprint *, int(d)\nend program t\n",
        ["20"]
    };

    // ── Write then READ roundtrip ────────────────────────────────────

    iio_roundtrip_negative_integer => {
        "program t\ncharacter(len=10) :: buf\ninteger :: x = -18, y\nwrite(buf, '(I0)') x\nread(buf, '(I0)') y\nprint *, y\nend program t\n",
        ["-18"]
    };

    iio_roundtrip_real_f62 => {
        "program t\ncharacter(len=10) :: buf\nreal :: x = 0.75, y\nwrite(buf, '(F6.2)') x\nread(buf, '(F6.2)') y\nprint *, y\nend program t\n",
        ["0.75"]
    };

    iio_roundtrip_logical_l1 => {
        "program t\ncharacter(len=4) :: buf\nlogical :: a = .true., b\nwrite(buf, '(L1)') a\nread(buf, '(L1)') b\nprint *, b\nend program t\n",
        ["true"]
    };

    iio_roundtrip_star_both_directions => {
        "program t\ncharacter(len=12) :: buf\ninteger :: x = 55, y\nwrite(buf, *) x\nread(buf, *) y\nprint *, y\nend program t\n",
        ["55"]
    };

    iio_roundtrip_string_a8 => {
        "program t\ncharacter(len=10) :: buf\ncharacter(len=4) :: s1 = 'vybe', s2\nwrite(buf, '(A4)') s1\nread(buf, '(A4)') s2\nprint *, s2\nend program t\n",
        ["vybe"]
    };

    // ── Parsing and utility patterns ─────────────────────────────────

    iio_len_trim_after_write_i0 => {
        "program t\ncharacter(len=10) :: buf\nwrite(buf, '(I0)') 404\nprint *, len_trim(buf)\nend program t\n",
        ["3"]
    };

    iio_parse_integers_in_do_loop => {
        "program t\ncharacter(len=4) :: vals(3) = ['1', '2', '3']\ninteger :: i, n, total\n total = 0\ndo i = 1, 3\nread(vals(i), '(I0)') n\ntotal = total + n\nend do\nprint *, total\nend program t\n",
        ["6"]
    };

    iio_internal_io_in_contained_subroutine => {
        "program t\ncall emit(9)\ncontains\nsubroutine emit(n)\ninteger, intent(in) :: n\ncharacter(len=6) :: buf\nwrite(buf, '(I0)') n\nprint *, trim(buf)\nend subroutine emit\nend program t\n",
        ["9"]
    };

    iio_iostat_zero_on_valid_read => {
        "program t\ncharacter(len=6) :: buf = '88'\ninteger :: n, ios\nread(buf, *, iostat=ios) n\nif (ios == 0) print *, n\nend program t\n",
        ["88"]
    };

    iio_two_sequential_reads_same_buffer => {
        "program t\ncharacter(len=10) :: buf = '2 3'\ninteger :: a, b\nread(buf, *) a\nread(buf, *) b\nprint *, a + b\nend program t\n",
        ["5"]
    };

    iio_build_key_equals_value_line => {
        "program t\ncharacter(len=16) :: buf\ninteger :: key = 7\nwrite(buf, '(A,I0)') 'id=', key\nprint *, trim(buf)\nend program t\n",
        ["id=7"]
    };

    iio_integer_arithmetic_after_parse => {
        "program t\ncharacter(len=6) :: buf = '6'\ninteger :: n\nread(buf, '(I0)') n\nprint *, n * n\nend program t\n",
        ["36"]
    };

    iio_real_compare_after_formatted_read => {
        "program t\ncharacter(len=10) :: buf = ' 1.00'\nreal :: x\nread(buf, '(F5.2)') x\nprint *, x == 1.0\nend program t\n",
        ["true"]
    };

    iio_write_read_index_verification => {
        "program t\ncharacter(len=12) :: buf\nwrite(buf, '(A)') 'needle'\nprint *, index(buf, 'eed')\nend program t\n",
        ["3"]
    };

    iio_list_read_trailing_spaces_ignored => {
        "program t\ncharacter(len=10) :: buf = '99    '\ninteger :: n\nread(buf, *) n\nprint *, n + 1\nend program t\n",
        ["100"]
    };
}
