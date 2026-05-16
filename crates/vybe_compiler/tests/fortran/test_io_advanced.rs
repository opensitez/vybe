use super::helpers::compile_ok;

// ── WRITE with FORMAT ─────────────────────────────────────────

#[test] fn write_fmt_integer() { compile_ok("program t\n  write(*, '(I5)') 42\nend program t\n"); }
#[test] fn write_fmt_real() { compile_ok("program t\n  write(*, '(F8.3)') 3.14159\nend program t\n"); }
#[test] fn write_fmt_string() { compile_ok("program t\n  write(*, '(A)') 'hello'\nend program t\n"); }
#[test] fn write_fmt_logical() { compile_ok("program t\n  write(*, '(L5)') .true.\nend program t\n"); }
#[test] fn write_fmt_multiple() { compile_ok("program t\n  integer :: n = 42\n  real :: x = 3.14\n  write(*, '(I5, F8.3)') n, x\nend program t\n"); }

#[test] fn write_fmt_scientific() { compile_ok("program t\n  write(*, '(E12.4)') 1.23456e10\nend program t\n"); }
#[test] fn write_fmt_general() { compile_ok("program t\n  write(*, '(G12.4)') 3.14\nend program t\n"); }
#[test] fn write_fmt_tab() { compile_ok("program t\n  write(*, '(A, 5X, A)') 'left', 'right'\nend program t\n"); }
#[test] fn write_fmt_newline() { compile_ok("program t\n  write(*, '(A, /, A)') 'line1', 'line2'\nend program t\n"); }
#[test] fn write_fmt_repeat() { compile_ok("program t\n  write(*, '(3I4)') 1, 2, 3\nend program t\n"); }
#[test] fn write_fmt_double() { compile_ok("program t\n  write(*, '(D20.12)') 3.141592653589793\nend program t\n"); }
#[test] fn write_fmt_binary() { compile_ok("program t\n  write(*, '(B8)') 255\nend program t\n"); }
#[test] fn write_fmt_octal() { compile_ok("program t\n  write(*, '(O8)') 255\nend program t\n"); }
#[test] fn write_fmt_hex() { compile_ok("program t\n  write(*, '(Z8)') 255\nend program t\n"); }

// ── Named FORMAT statements ───────────────────────────────────

#[test] fn format_label() {
    compile_ok(r#"
program test
    integer :: i = 7
    write(*, 100) i
100 format(I5)
end program test
"#);
}

#[test] fn format_label_float() {
    compile_ok(r#"
program test
    real :: x = 2.718
200 format(F10.4)
    write(*, 200) x
end program test
"#);
}

#[test] fn format_label_string() {
    compile_ok(r#"
program test
    character(len=5) :: s = 'hello'
300 format(A10)
    write(*, 300) s
end program test
"#);
}

// ── READ ─────────────────────────────────────────────────────

#[test] fn read_list_directed() { compile_ok("program t\n  integer :: n\n  read(*, *) n\n  print *, n\nend program t\n"); }
#[test] fn read_formatted() { compile_ok("program t\n  integer :: n\n  read(*, '(I5)') n\n  print *, n\nend program t\n"); }
#[test] fn read_multiple() { compile_ok("program t\n  integer :: a, b\n  read(*, *) a, b\n  print *, a + b\nend program t\n"); }
#[test] fn read_string() { compile_ok("program t\n  character(len=20) :: s\n  read(*, *) s\n  print *, s\nend program t\n"); }

// ── File I/O — OPEN / CLOSE ───────────────────────────────────

#[test] fn open_close_basic() {
    compile_ok(r#"
program test
    integer :: unit = 10
    open(unit=10, file='test.txt', status='replace')
    write(10, *) 'hello'
    close(10)
end program test
"#);
}

#[test] fn open_read_write() {
    compile_ok(r#"
program test
    open(unit=20, file='data.txt', status='replace', action='write')
    write(20, '(A)') 'test data'
    close(20)
end program test
"#);
}

#[test] fn open_status_old() {
    compile_ok(r#"
program test
    open(unit=30, file='existing.txt', status='old', action='read', iostat=ios)
    integer :: ios
    if (ios /= 0) print *, 'file not found'
end program test
"#);
}

#[test] fn open_status_scratch() {
    compile_ok(r#"
program test
    open(unit=40, status='scratch')
    write(40, *) 42
    rewind(40)
    close(40, status='delete')
end program test
"#);
}

#[test] fn open_newunit() {
    compile_ok(r#"
program test
    integer :: u
    open(newunit=u, file='tmp.txt', status='replace')
    write(u, *) 'newunit test'
    close(u)
end program test
"#);
}

// ── REWIND / BACKSPACE / ENDFILE ──────────────────────────────

#[test] fn rewind_stmt() {
    compile_ok(r#"
program test
    open(unit=10, status='scratch')
    write(10, *) 1, 2, 3
    rewind(10)
    close(10)
end program test
"#);
}

#[test] fn backspace_stmt() {
    compile_ok(r#"
program test
    open(unit=10, status='scratch')
    write(10, *) 1
    write(10, *) 2
    backspace(10)
    close(10)
end program test
"#);
}

#[test] fn endfile_stmt() {
    compile_ok(r#"
program test
    open(unit=10, status='scratch')
    write(10, *) 42
    endfile(10)
    close(10)
end program test
"#);
}

// ── INQUIRE ───────────────────────────────────────────────────

#[test] fn inquire_exist() {
    compile_ok(r#"
program test
    logical :: exists
    inquire(file='test.txt', exist=exists)
    print *, exists
end program test
"#);
}

#[test] fn inquire_unit() {
    compile_ok(r#"
program test
    logical :: opened
    inquire(unit=10, opened=opened)
    print *, opened
end program test
"#);
}

#[test] fn inquire_size() {
    compile_ok(r#"
program test
    integer :: fsize
    inquire(file='test.txt', size=fsize)
    print *, fsize
end program test
"#);
}

#[test] fn inquire_name() {
    compile_ok(r#"
program test
    character(len=100) :: fname
    open(unit=10, status='scratch')
    inquire(unit=10, name=fname)
    close(10)
    print *, 'ok'
end program test
"#);
}

// ── FLUSH (Fortran 2003) ──────────────────────────────────────

#[test] fn flush_stmt() {
    compile_ok(r#"
program test
    open(unit=10, file='out.txt', status='replace')
    write(10, *) 'buffered data'
    flush(10)
    close(10)
end program test
"#);
}

#[test] fn flush_stdout() {
    compile_ok(r#"
program test
    print *, 'about to flush'
    flush(6)
end program test
"#);
}

// ── IOSTAT and ERR ────────────────────────────────────────────

#[test] fn read_iostat() {
    compile_ok(r#"
program test
    integer :: n, ios
    read(*, *, iostat=ios) n
    if (ios /= 0) then
        print *, 'read error'
    else
        print *, n
    end if
end program test
"#);
}

#[test] fn open_iostat() {
    compile_ok(r#"
program test
    integer :: ios
    open(unit=10, file='nosuchfile.txt', status='old', iostat=ios)
    if (ios /= 0) print *, 'could not open'
end program test
"#);
}

#[test] fn write_err() {
    compile_ok(r#"
program test
    integer :: ios
    write(*, *, iostat=ios) 42
    if (ios /= 0) print *, 'write error'
end program test
"#);
}

// ── Stream I/O (Fortran 2003) ─────────────────────────────────

#[test] fn stream_write() {
    compile_ok(r#"
program test
    open(unit=10, file='stream.bin', access='stream', form='unformatted', &
         status='replace')
    write(10) 42
    close(10)
end program test
"#);
}

#[test] fn stream_read_write() {
    compile_ok(r#"
program test
    integer :: x, y
    open(unit=10, file='stream.bin', access='stream', form='unformatted', &
         status='replace')
    write(10) 100, 200
    rewind(10)
    read(10) x, y
    close(10)
    print *, x, y
end program test
"#);
}

// ── NAMELIST ─────────────────────────────────────────────────

#[test] fn namelist_write() {
    compile_ok(r#"
program test
    integer :: x = 10, y = 20
    real :: z = 3.14
    namelist /cfg/ x, y, z
    write(*, nml=cfg)
end program test
"#);
}

#[test] fn namelist_read() {
    compile_ok(r#"
program test
    integer :: nx = 10, ny = 20
    namelist /grid/ nx, ny
    read(*, nml=grid, iostat=ios)
    integer :: ios
    print *, nx
end program test
"#);
}

// ── ADVANCE='NO' (non-advancing I/O) ─────────────────────────

#[test] fn non_advancing_write() {
    compile_ok(r#"
program test
    write(*, '(A)', advance='no') 'no newline'
    write(*, '(A)') ' here'
end program test
"#);
}

// ── Unformatted I/O ───────────────────────────────────────────

#[test] fn unformatted_write() {
    compile_ok(r#"
program test
    open(unit=10, file='bin.dat', form='unformatted', status='replace')
    write(10) 42, 3.14, .true.
    close(10)
end program test
"#);
}

#[test] fn unformatted_read() {
    compile_ok(r#"
program test
    integer :: n
    open(unit=10, file='bin.dat', form='unformatted', status='replace')
    write(10) 99
    rewind(10)
    read(10) n
    close(10)
    print *, n
end program test
"#);
}

// ── PRINT with formatting ──────────────────────────────────────

#[test] fn print_fmt_integer() { compile_ok("program t\n  print '(I8)', 12345\nend program t\n"); }
#[test] fn print_fmt_real() { compile_ok("program t\n  print '(F10.4)', 3.14159\nend program t\n"); }
#[test] fn print_fmt_string() { compile_ok("program t\n  print '(A)', 'hello world'\nend program t\n"); }
