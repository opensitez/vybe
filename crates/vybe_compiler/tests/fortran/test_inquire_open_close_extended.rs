//! Extended OPEN/CLOSE/INQUIRE: newunit, status, position, iostat, and unit
//! lifecycle with runtime assertions. Distinct from `test_io_advanced.rs`
//! compile smokes, `test_io_file_position.rs` rewind/backspace focus, and
//! `test_internal_io_extended.rs` list-directed internal I/O.

use super::helpers::compile_ok;

fortran_cases! {
    // ── Scratch unit lifecycle (10) ───────────────────────────────────

    ioc_scratch_write_read_integer => {
        "program t\ninteger :: n\nopen(20, status='scratch')\nwrite(20, '(I0)') 123\nrewind(20)\nread(20, '(I0)') n\nclose(20)\nprint *, n\nend program t\n",
        ["123"]
    };
    ioc_scratch_two_values_sum => {
        "program t\ninteger :: a, b\nopen(21, status='scratch')\nwrite(21, *) 10, 20\nrewind(21)\nread(21, *) a, b\nclose(21)\nprint *, a + b\nend program t\n",
        ["30"]
    };
    ioc_scratch_close_reopen_new_data => {
        "program t\ninteger :: v\nopen(22, status='scratch')\nwrite(22, '(I0)') 7\nclose(22)\nopen(22, status='scratch')\nwrite(22, '(I0)') 42\nrewind(22)\nread(22, '(I0)') v\nclose(22)\nprint *, v\nend program t\n",
        ["42"]
    };
    ioc_scratch_formatted_three_lines => {
        "program t\ninteger :: x, y, z\nopen(23, status='scratch')\nwrite(23, '(I0)') 1\nwrite(23, '(I0)') 2\nwrite(23, '(I0)') 3\nrewind(23)\nread(23, '(I0)') x\nread(23, '(I0)') y\nread(23, '(I0)') z\nclose(23)\nprint *, x\nprint *, y\nprint *, z\nend program t\n",
        ["1", "2", "3"]
    };
    ioc_scratch_real_value_roundtrip => {
        "program t\nreal :: r\nopen(24, status='scratch')\nwrite(24, '(F0.1)') 3.5\nrewind(24)\nread(24, '(F0.1)') r\nclose(24)\nprint *, int(r * 10)\nend program t\n",
        ["35"]
    };
    ioc_dual_scratch_independent => {
        "program t\ninteger :: a, b\nopen(30, status='scratch')\nopen(31, status='scratch')\nwrite(30, '(I0)') 11\nwrite(31, '(I0)') 22\nrewind(30)\nrewind(31)\nread(30, '(I0)') a\nread(31, '(I0)') b\nclose(30)\nclose(31)\nprint *, a\nprint *, b\nend program t\n",
        ["11", "22"]
    };
    ioc_scratch_logical_roundtrip => {
        "program t\nlogical :: flag\nopen(32, status='scratch')\nwrite(32, *) .true.\nrewind(32)\nread(32, *) flag\nclose(32)\nif (flag) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_scratch_character_roundtrip => {
        "program t\ncharacter(len=3) :: s\nopen(33, status='scratch')\nwrite(33, '(A)') 'abc'\nrewind(33)\nread(33, '(A)') s\nclose(33)\nprint *, s\nend program t\n",
        ["abc"]
    };
    ioc_scratch_multiple_rewind_same_value => {
        "program t\ninteger :: n\nopen(34, status='scratch')\nwrite(34, '(I0)') 55\nrewind(34)\nread(34, '(I0)') n\nrewind(34)\nread(34, '(I0)') n\nclose(34)\nprint *, n\nend program t\n",
        ["55"]
    };
    ioc_scratch_empty_then_write => {
        "program t\ninteger :: n\nopen(35, status='scratch')\nwrite(35, '(I0)') 0\nrewind(35)\nread(35, '(I0)') n\nclose(35)\nprint *, n\nend program t\n",
        ["0"]
    };

    // ── Named file status='replace' (8) ───────────────────────────────

    ioc_replace_create_read_back => {
        "program t\ninteger :: n\nopen(40, file='ioc_ext_rep1.dat', status='replace')\nwrite(40, '(I0)') 999\nrewind(40)\nread(40, '(I0)') n\nclose(40, status='delete')\nprint *, n\nend program t\n",
        ["999"]
    };
    ioc_replace_truncates_old_content => {
        "program t\ninteger :: n\nopen(41, file='ioc_ext_rep2.dat', status='replace')\nwrite(41, '(I0)') 111\nclose(41)\nopen(41, file='ioc_ext_rep2.dat', status='replace')\nwrite(41, '(I0)') 222\nrewind(41)\nread(41, '(I0)') n\nclose(41, status='delete')\nprint *, n\nend program t\n",
        ["222"]
    };
    ioc_replace_two_integers => {
        "program t\ninteger :: a, b\nopen(42, file='ioc_ext_rep3.dat', status='replace')\nwrite(42, *) 3, 4\nrewind(42)\nread(42, *) a, b\nclose(42, status='delete')\nprint *, a * b\nend program t\n",
        ["12"]
    };
    ioc_replace_formatted_record => {
        "program t\ninteger :: n\nopen(43, file='ioc_ext_rep4.dat', status='replace', form='formatted')\nwrite(43, '(I0)') 77\nrewind(43)\nread(43, '(I0)') n\nclose(43, status='delete')\nprint *, n\nend program t\n",
        ["77"]
    };
    ioc_replace_action_readwrite => {
        "program t\ninteger :: n\nopen(44, file='ioc_ext_rep5.dat', status='replace', action='readwrite')\nwrite(44, '(I0)') 88\nrewind(44)\nread(44, '(I0)') n\nclose(44, status='delete')\nprint *, n\nend program t\n",
        ["88"]
    };
    ioc_replace_sequential_access => {
        "program t\ninteger :: n\nopen(45, file='ioc_ext_rep6.dat', status='replace', access='sequential')\nwrite(45, '(I0)') 66\nrewind(45)\nread(45, '(I0)') n\nclose(45, status='delete')\nprint *, n\nend program t\n",
        ["66"]
    };
    ioc_replace_append_after_reopen => {
        "program t\ninteger :: n\nopen(46, file='ioc_ext_rep7.dat', status='replace')\nwrite(46, '(I0)') 10\nclose(46)\nopen(46, file='ioc_ext_rep7.dat', status='old', position='append')\nwrite(46, '(I0)') 5\nrewind(46)\nread(46, '(I0)') n\nclose(46, status='delete')\nprint *, n\nend program t\n",
        ["10"]
    };
    ioc_replace_delete_on_close => {
        "program t\ninteger :: n\nopen(47, file='ioc_ext_rep8.dat', status='replace')\nwrite(47, '(I0)') 44\nclose(47, status='delete')\nprint *, 44\nend program t\n",
        ["44"]
    };

    // ── INQUIRE runtime by unit (10) ──────────────────────────────────

    ioc_inquire_opened_true_on_scratch => {
        "program t\nlogical :: opened\nopen(50, status='scratch')\ninquire(unit=50, opened=opened)\nclose(50)\nif (opened) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_inquire_opened_false_before_open => {
        "program t\nlogical :: opened\ninquire(unit=99, opened=opened)\nif (opened) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["0"]
    };
    ioc_inquire_opened_false_after_close => {
        "program t\nlogical :: opened\nopen(51, status='scratch')\nclose(51)\ninquire(unit=51, opened=opened)\nif (opened) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["0"]
    };
    ioc_inquire_number_returns_unit => {
        "program t\ninteger :: num\nopen(52, status='scratch')\ninquire(unit=52, number=num)\nclose(52)\nprint *, num\nend program t\n",
        ["52"]
    };
    ioc_inquire_named_file_exists_after_replace => {
        "program t\nlogical :: exists\nopen(53, file='ioc_ext_exist.dat', status='replace')\nwrite(53, '(I0)') 1\nclose(53)\ninquire(file='ioc_ext_exist.dat', exist=exists)\nopen(53, file='ioc_ext_exist.dat', status='old')\nclose(53, status='delete')\nif (exists) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_inquire_named_file_not_exist => {
        "program t\nlogical :: exists\ninquire(file='ioc_ext_no_such_file.dat', exist=exists)\nif (exists) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["0"]
    };
    ioc_inquire_sequential_form_on_scratch => {
        "program t\ncharacter(len=20) :: acc, frm\nopen(54, status='scratch')\ninquire(unit=54, access=acc, form=frm)\nclose(54)\nprint *, acc(1:10)\nend program t\n",
        ["SEQUENTIAL"]
    };
    ioc_inquire_formatted_on_replace_file => {
        "program t\ncharacter(len=20) :: frm\nopen(55, file='ioc_ext_form.dat', status='replace', form='formatted')\nwrite(55, '(I0)') 1\ninquire(unit=55, form=frm)\nclose(55, status='delete')\nprint *, frm(1:9)\nend program t\n",
        ["FORMATTED"]
    };
    ioc_inquire_iostat_zero_on_success_open => {
        "program t\ninteger :: ios\nopen(56, file='ioc_ext_ios.dat', status='replace', iostat=ios)\nclose(56, status='delete')\nprint *, ios\nend program t\n",
        ["0"]
    };
    ioc_inquire_size_after_write => {
        "program t\ninteger :: sz\nopen(57, file='ioc_ext_size.dat', status='replace')\nwrite(57, '(I0)') 12345\ninquire(unit=57, size=sz)\nclose(57, status='delete')\nif (sz > 0) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };

    // ── newunit assignment (6) ──────────────────────────────────────────

    ioc_newunit_write_read_integer => {
        "program t\ninteger :: u, n\nopen(newunit=u, file='ioc_ext_new1.dat', status='replace')\nwrite(u, '(I0)') 314\nrewind(u)\nread(u, '(I0)') n\nclose(u, status='delete')\nprint *, n\nend program t\n",
        ["314"]
    };
    ioc_newunit_two_handles => {
        "program t\ninteger :: u1, u2, a, b\nopen(newunit=u1, file='ioc_ext_new2a.dat', status='replace')\nopen(newunit=u2, file='ioc_ext_new2b.dat', status='replace')\nwrite(u1, '(I0)') 10\nwrite(u2, '(I0)') 20\nrewind(u1)\nrewind(u2)\nread(u1, '(I0)') a\nread(u2, '(I0)') b\nclose(u1, status='delete')\nclose(u2, status='delete')\nprint *, a + b\nend program t\n",
        ["30"]
    };
    ioc_newunit_positive_unit_number => {
        "program t\ninteger :: u\nopen(newunit=u, file='ioc_ext_new3.dat', status='replace')\nclose(u, status='delete')\nif (u > 0) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_newunit_list_directed => {
        "program t\ninteger :: u, a, b\nopen(newunit=u, file='ioc_ext_new4.dat', status='replace')\nwrite(u, *) 6, 7\nrewind(u)\nread(u, *) a, b\nclose(u, status='delete')\nprint *, a + b\nend program t\n",
        ["13"]
    };
    ioc_newunit_real_roundtrip => {
        "program t\ninteger :: u\nreal :: r\nopen(newunit=u, file='ioc_ext_new5.dat', status='replace')\nwrite(u, '(F0.1)') 2.5\nrewind(u)\nread(u, '(F0.1)') r\nclose(u, status='delete')\nprint *, int(r * 10)\nend program t\n",
        ["25"]
    };
    ioc_newunit_inquire_opened => {
        "program t\ninteger :: u\nlogical :: opened\nopen(newunit=u, file='ioc_ext_new6.dat', status='replace')\ninquire(unit=u, opened=opened)\nclose(u, status='delete')\nif (opened) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };

    // ── IOSTAT on read/write success and failure (8) ──────────────────

    ioc_iostat_zero_on_formatted_read => {
        "program t\ninteger :: n, ios\nopen(60, status='scratch')\nwrite(60, '(I0)') 42\nrewind(60)\nread(60, '(I0)', iostat=ios) n\nclose(60)\nprint *, ios\nprint *, n\nend program t\n",
        ["0", "42"]
    };
    ioc_iostat_zero_on_list_write => {
        "program t\ninteger :: ios\nopen(61, status='scratch')\nwrite(61, *, iostat=ios) 1, 2, 3\nclose(61)\nprint *, ios\nend program t\n",
        ["0"]
    };
    ioc_iostat_endfile_detection => {
        "program t\ninteger :: n, ios\nopen(62, status='scratch')\nwrite(62, '(I0)') 1\nrewind(62)\nread(62, '(I0)', iostat=ios) n\nread(62, '(I0)', iostat=ios) n\nclose(62)\nprint *, ios\nend program t\n",
        ["-1"]
    };
    ioc_iostat_open_old_missing_file => {
        "program t\ninteger :: ios\nopen(63, file='ioc_ext_missing.dat', status='old', iostat=ios)\nif (ios /= 0) then\nprint *, 1\nelse\nclose(63)\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_iostat_close_already_closed_unit => {
        "program t\ninteger :: ios\nopen(64, status='scratch')\nclose(64)\nclose(64, iostat=ios)\nprint *, ios\nend program t\n",
        ["0"]
    };
    ioc_iostat_read_from_empty_scratch => {
        "program t\ninteger :: n, ios\nopen(65, status='scratch')\nread(65, '(I0)', iostat=ios) n\nclose(65)\nif (ios /= 0) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_iostat_write_after_close_fails => {
        "program t\ninteger :: ios\nopen(66, status='scratch')\nclose(66)\nwrite(66, '(I0)', iostat=ios) 1\nif (ios /= 0) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
    ioc_iostat_inquire_closed_unit => {
        "program t\ninteger :: ios\nlogical :: opened\nopen(67, status='scratch')\nclose(67)\ninquire(unit=67, opened=opened, iostat=ios)\nprint *, ios\nif (opened) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["0", "0"]
    };

    // ── CLOSE variants and position= (8) ──────────────────────────────

    ioc_close_status_keep_file => {
        "program t\ninteger :: n\nopen(70, file='ioc_ext_keep.dat', status='replace')\nwrite(70, '(I0)') 50\nclose(70, status='keep')\nopen(70, file='ioc_ext_keep.dat', status='old')\nread(70, '(I0)') n\nclose(70, status='delete')\nprint *, n\nend program t\n",
        ["50"]
    };
    ioc_close_status_delete => {
        "program t\nopen(71, file='ioc_ext_del.dat', status='replace')\nwrite(71, '(I0)') 1\nclose(71, status='delete')\nprint *, 1\nend program t\n",
        ["1"]
    };
    ioc_position_rewind_after_append => {
        "program t\ninteger :: n\nopen(72, file='ioc_ext_pos.dat', status='replace')\nwrite(72, '(I0)') 100\nclose(72)\nopen(72, file='ioc_ext_pos.dat', status='old', position='append')\nwrite(72, '(I0)') 200\nrewind(72)\nread(72, '(I0)') n\nclose(72, status='delete')\nprint *, n\nend program t\n",
        ["100"]
    };
    ioc_position_asis_reopen => {
        "program t\ninteger :: n\nopen(73, file='ioc_ext_asis.dat', status='replace')\nwrite(73, '(I0)') 33\nclose(73)\nopen(73, file='ioc_ext_asis.dat', status='old', position='asis')\nread(73, '(I0)') n\nclose(73, status='delete')\nprint *, n\nend program t\n",
        ["33"]
    };
    ioc_unformatted_scratch_roundtrip => {
        "program t\ninteger :: a, b\nopen(74, status='scratch', form='unformatted')\nwrite(74) 8, 9\nrewind(74)\nread(74) a, b\nclose(74)\nprint *, a + b\nend program t\n",
        ["17"]
    };
    ioc_stream_unformatted_roundtrip => {
        "program t\ninteger :: v\nopen(75, file='ioc_ext_stream.dat', access='stream', form='unformatted', status='replace')\nwrite(75) 64\nrewind(75)\nread(75) v\nclose(75, status='delete')\nprint *, v\nend program t\n",
        ["64"]
    };
    ioc_close_unit_without_status => {
        "program t\ninteger :: n\nopen(76, file='ioc_ext_plain.dat', status='replace')\nwrite(76, '(I0)') 17\nclose(76)\nopen(76, file='ioc_ext_plain.dat', status='old')\nread(76, '(I0)') n\nclose(76, status='delete')\nprint *, n\nend program t\n",
        ["17"]
    };
    ioc_multiple_close_same_file_different_units => {
        "program t\ninteger :: n\nopen(77, file='ioc_ext_multi.dat', status='replace')\nwrite(77, '(I0)') 91\nclose(77)\nopen(78, file='ioc_ext_multi.dat', status='old')\nread(78, '(I0)') n\nclose(78, status='delete')\nprint *, n\nend program t\n",
        ["91"]
    };

    // ── Internal I/O fallback for inquire-like patterns (5) ───────────

    ioc_internal_write_read_integer => {
        "program t\ncharacter(len=20) :: buf\ninteger :: n\nwrite(buf, '(I0)') 456\nread(buf, '(I0)') n\nprint *, n\nend program t\n",
        ["456"]
    };
    ioc_internal_write_read_two_values => {
        "program t\ncharacter(len=30) :: buf\ninteger :: a, b\nwrite(buf, *) 12, 34\nread(buf, *) a, b\nprint *, a\nprint *, b\nend program t\n",
        ["12", "34"]
    };
    ioc_internal_formatted_roundtrip => {
        "program t\ncharacter(len=10) :: buf\nreal :: r\nwrite(buf, '(F0.1)') 4.5\nread(buf, '(F0.1)') r\nprint *, int(r * 10)\nend program t\n",
        ["45"]
    };
    ioc_internal_character_preservation => {
        "program t\ncharacter(len=10) :: buf, s\ns = 'hello'\nwrite(buf, '(A)') s\nread(buf, '(A)') s\nprint *, s\nend program t\n",
        ["hello"]
    };
    ioc_internal_iostat_on_bad_read => {
        "program t\ncharacter(len=5) :: buf\ninteger :: n, ios\nbuf = 'abc'\nread(buf, '(I0)', iostat=ios) n\nif (ios /= 0) then\nprint *, 1\nelse\nprint *, 0\nend if\nend program t\n",
        ["1"]
    };
}

// ── Compile-only: OPEN/INQUIRE forms not requiring runtime files ──────

#[test]
fn ioc_compile_open_access_direct() {
    compile_ok(
        r#"
program t
    open(10, file='ioc_ext_direct.dat', access='direct', recl=10, status='replace')
    write(10, rec=1) 'test'
    close(10, status='delete')
    print *, 1
end program t
"#,
    );
}

#[test]
fn ioc_compile_inquire_all_attributes() {
    compile_ok(
        r#"
program t
    logical :: opened, named, exist
    integer :: num, rec, ios
    character(len=20) :: acc, frm, nm
    open(10, file='ioc_ext_inq.dat', status='replace')
    inquire(unit=10, opened=opened, number=num, named=named, name=nm, &
            access=acc, form=frm, rec=rec, iostat=ios)
    close(10, status='delete')
    print *, 1
end program t
"#,
    );
}

#[test]
fn ioc_compile_open_pad_yes() {
    compile_ok(
        r#"
program t
    open(10, file='ioc_ext_pad.dat', status='replace', pad='yes')
    write(10, '(A)') 'x'
    close(10, status='delete')
    print *, 1
end program t
"#,
    );
}

#[test]
fn ioc_compile_open_delim_quote() {
    compile_ok(
        r#"
program t
    open(10, file='ioc_ext_delim.dat', status='replace', delim='quote')
    write(10, *) 'a'
    close(10, status='delete')
    print *, 1
end program t
"#,
    );
}

#[test]
fn ioc_compile_inquire_by_filename_only() {
    compile_ok(
        r#"
program t
    logical :: opened
    inquire(file='ioc_ext_fn.dat', opened=opened)
    print *, 0
end program t
"#,
    );
}

#[test]
fn ioc_compile_open_convert_native() {
    compile_ok(
        r#"
program t
    open(10, file='ioc_ext_conv.dat', status='replace', convert='native')
    write(10, '(I0)') 1
    close(10, status='delete')
    print *, 1
end program t
"#,
    );
}

#[test]
fn ioc_compile_close_err_label() {
    compile_ok(
        r#"
program t
    integer :: ios
    open(10, status='scratch')
    close(10, iostat=ios)
    print *, ios
end program t
"#,
    );
}
