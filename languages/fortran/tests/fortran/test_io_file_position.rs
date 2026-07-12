//! External file I/O positioning: scratch open/close, REWIND, BACKSPACE,
//! INQUIRE by unit/file, stream `position=`, `status='replace'`, and `iostat`
//! on read failure. Distinct from `test_io.rs`, `test_io_advanced.rs`, and
//! `test_internal_io_extended.rs`.

use super::helpers::compile_ok;

fortran_cases! {
    // ── Scratch open/close with positioned readback ─────────────────

    fio_scratch_list_directed_sum_after_rewind => {
        "program t\ninteger :: a, b\nopen(10, status='scratch')\nwrite(10, *) 10, 20\nrewind(10)\nread(10, *) a, b\nclose(10)\nprint *, a + b\nend program t\n",
        ["30"]
    };

    fio_scratch_formatted_i0_roundtrip => {
        "program t\ninteger :: n\nopen(11, status='scratch')\nwrite(11, '(I0)') 55\nrewind(11)\nread(11, '(I0)') n\nclose(11)\nprint *, n\nend program t\n",
        ["55"]
    };

    fio_scratch_close_then_reopen_rewind_read => {
        "program t\ninteger :: v\nopen(12, status='scratch')\nwrite(12, '(I0)') 9\nclose(12)\nopen(12, status='scratch')\nwrite(12, '(I0)') 4\nrewind(12)\nread(12, '(I0)') v\nclose(12)\nprint *, v\nend program t\n",
        ["4"]
    };

    fio_dual_scratch_units_independent_values => {
        "program t\ninteger :: a, b\nopen(15, status='scratch')\nopen(16, status='scratch')\nwrite(15, '(I0)') 3\nwrite(16, '(I0)') 7\nrewind(15)\nrewind(16)\nread(15, '(I0)') a\nread(16, '(I0)') b\nclose(15)\nclose(16)\nprint *, a\nprint *, b\nend program t\n",
        ["3", "7"]
    };

    // ── REWIND across formatted records ───────────────────────────────

    fio_rewind_reread_two_formatted_records => {
        "program t\ninteger :: a, b\nopen(20, status='scratch')\nwrite(20, '(I0)') 1\nwrite(20, '(I0)') 2\nrewind(20)\nread(20, '(I0)') a\nread(20, '(I0)') b\nclose(20)\nprint *, a\nprint *, b\nend program t\n",
        ["1", "2"]
    };

    fio_rewind_after_endfile_then_append => {
        "program t\ninteger :: tail\nopen(21, status='scratch')\nwrite(21, '(I0)') 5\nendfile(21)\nrewind(21)\nwrite(21, '(I0)') 6\nrewind(21)\nread(21, '(I0)') tail\nclose(21)\nprint *, tail\nend program t\n",
        ["6"]
    };

    fio_rewind_overwrites_first_record_value => {
        "program t\ninteger :: first\nopen(22, status='scratch')\nwrite(22, '(I0)') 100\nwrite(22, '(I0)') 200\nrewind(22)\nread(22, '(I0)') first\nwrite(22, '(I0)') 300\nrewind(22)\nread(22, '(I0)') first\nclose(22)\nprint *, first\nend program t\n",
        ["300"]
    };

    // ── status='replace' truncation semantics ─────────────────────────

    fio_replace_single_integer_roundtrip => {
        "program t\ninteger :: n\nopen(30, file='fio_replace_one.dat', status='replace')\nwrite(30, '(I0)') 88\nrewind(30)\nread(30, '(I0)') n\nclose(30, status='delete')\nprint *, n\nend program t\n",
        ["88"]
    };

    fio_replace_discards_prior_session_content => {
        "program t\ninteger :: n\nopen(31, file='fio_replace_stale.dat', status='replace')\nwrite(31, '(I0)') 111\nclose(31)\nopen(31, file='fio_replace_stale.dat', status='replace')\nwrite(31, '(I0)') 222\nrewind(31)\nread(31, '(I0)') n\nclose(31, status='delete')\nprint *, n\nend program t\n",
        ["222"]
    };

    // ── Stream access with REWIND readback ────────────────────────────

    fio_stream_scratch_rewind_product => {
        "program t\ninteger :: a, b\nopen(40, status='scratch', access='stream', form='unformatted')\nwrite(40) 8, 9\nrewind(40)\nread(40) a, b\nclose(40)\nprint *, a * b\nend program t\n",
        ["72"]
    };

    fio_stream_replace_file_rewind_read => {
        "program t\ninteger :: v\nopen(41, file='fio_stream_pos.dat', access='stream', form='unformatted', status='replace')\nwrite(41) 64\nrewind(41)\nread(41) v\nclose(41, status='delete')\nprint *, v\nend program t\n",
        ["64"]
    };
}

// ── Scratch close variants ────────────────────────────────────────────

#[test]
fn fio_scratch_close_status_delete() {
    compile_ok(
        r#"
program t
    open(10, status='scratch')
    write(10, '(I0)') 1
    close(10, status='delete')
end program t
"#,
    );
}

// ── BACKSPACE positioning ─────────────────────────────────────────────

#[test]
fn fio_backspace_after_three_formatted_records() {
    compile_ok(
        r#"
program t
    open(10, status='scratch')
    write(10, '(I0)') 1
    write(10, '(I0)') 2
    write(10, '(I0)') 3
    backspace(10)
    close(10)
end program t
"#,
    );
}

#[test]
fn fio_backspace_twice_then_rewrite_last_line() {
    compile_ok(
        r#"
program t
    open(10, status='scratch')
    write(10, '(A)') 'alpha'
    write(10, '(A)') 'beta'
    write(10, '(A)') 'gamma'
    backspace(10)
    backspace(10)
    write(10, '(A)') 'delta'
    close(10)
end program t
"#,
    );
}

// ── INQUIRE by unit and by file ───────────────────────────────────────

#[test]
fn fio_inquire_unit_opened_after_scratch_open() {
    compile_ok(
        r#"
program t
    logical :: opened
    open(10, status='scratch')
    inquire(unit=10, opened=opened)
    close(10)
    print *, opened
end program t
"#,
    );
}

#[test]
fn fio_inquire_file_exist_after_replace_create() {
    compile_ok(
        r#"
program t
    logical :: exists
    open(10, file='fio_inquire_exist.dat', status='replace')
    write(10, '(I0)') 1
    close(10)
    inquire(file='fio_inquire_exist.dat', exist=exists)
    print *, exists
end program t
"#,
    );
}

#[test]
fn fio_inquire_unit_access_and_form_sequential() {
    compile_ok(
        r#"
program t
    character(len=16) :: access, form
    open(10, file='fio_inquire_attrs.dat', status='replace')
    inquire(unit=10, access=access, form=form)
    close(10, status='delete')
    print *, trim(access)
    print *, trim(form)
end program t
"#,
    );
}

#[test]
fn fio_inquire_stream_pos_after_unformatted_write() {
    compile_ok(
        r#"
program t
    integer :: pos
    open(10, status='scratch', access='stream', form='unformatted')
    write(10) 1, 2, 3
    inquire(unit=10, pos=pos)
    close(10)
    print *, pos
end program t
"#,
    );
}

// ── Stream `position=` on OPEN ────────────────────────────────────────

#[test]
fn fio_open_stream_position_rewind_specifier() {
    compile_ok(
        r#"
program t
    integer :: v
    open(10, file='fio_open_rewind.dat', access='stream', form='unformatted', &
         status='replace', position='rewind')
    write(10) 12
    rewind(10)
    read(10) v
    close(10, status='delete')
    print *, v
end program t
"#,
    );
}

#[test]
fn fio_open_stream_position_append_specifier() {
    compile_ok(
        r#"
program t
    open(10, file='fio_open_append.dat', access='stream', form='unformatted', &
         status='replace', position='append')
    write(10) 1
    close(10, status='delete')
end program t
"#,
    );
}

// ── `iostat` on file read failure ─────────────────────────────────────

#[test]
fn fio_read_iostat_eof_on_second_scratch_read() {
    compile_ok(
        r#"
program t
    integer :: n, ios
    open(10, status='scratch')
    write(10, '(I0)') 42
    rewind(10)
    read(10, *, iostat=ios) n
    read(10, *, iostat=ios) n
    if (ios /= 0) print *, 1
    close(10)
end program t
"#,
    );
}

#[test]
fn fio_read_iostat_short_formatted_record() {
    compile_ok(
        r#"
program t
    integer :: a, b, ios
    open(10, status='scratch')
    write(10, '(I0)') 7
    rewind(10)
    read(10, '(2I4)', iostat=ios) a, b
    if (ios /= 0) print *, ios
    close(10)
end program t
"#,
    );
}
