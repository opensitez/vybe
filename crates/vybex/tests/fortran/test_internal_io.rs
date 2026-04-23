use super::helpers::compile_ok;

// ── WRITE to character variable (internal file) ───────────────

#[test] fn write_int_to_string() {
    compile_ok(r#"
program test
    character(len=20) :: buf
    write(buf, *) 42
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_real_to_string() {
    compile_ok(r#"
program test
    character(len=30) :: buf
    write(buf, *) 3.14159
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_formatted_int() {
    compile_ok(r#"
program test
    character(len=10) :: buf
    write(buf, '(I5)') 42
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_formatted_real() {
    compile_ok(r#"
program test
    character(len=15) :: buf
    write(buf, '(F8.3)') 3.14159
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_multiple_values() {
    compile_ok(r#"
program test
    character(len=30) :: buf
    write(buf, *) 1, 2, 3
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_logical_to_string() {
    compile_ok(r#"
program test
    character(len=5) :: buf
    write(buf, '(L5)') .true.
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_string_to_string() {
    compile_ok(r#"
program test
    character(len=20) :: buf
    write(buf, '(A)') 'hello'
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_mixed_format() {
    compile_ok(r#"
program test
    character(len=30) :: buf
    integer :: n = 10
    real :: x = 3.14
    write(buf, '(I4, F8.3)') n, x
    print *, trim(buf)
end program test
"#);
}

#[test] fn write_double_to_string() {
    compile_ok(r#"
program test
    character(len=30) :: buf
    real(kind=8) :: d = 2.718281828d0
    write(buf, '(D20.12)') d
    print *, trim(buf)
end program test
"#);
}

// ── READ from character variable (internal file) ──────────────

#[test] fn read_int_from_string() {
    compile_ok(r#"
program test
    character(len=10) :: buf = '   42'
    integer :: n
    read(buf, *) n
    print *, n
end program test
"#);
}

#[test] fn read_real_from_string() {
    compile_ok(r#"
program test
    character(len=15) :: buf = '  3.14'
    real :: x
    read(buf, *) x
    print *, x
end program test
"#);
}

#[test] fn read_formatted_int() {
    compile_ok(r#"
program test
    character(len=10) :: buf = '   42'
    integer :: n
    read(buf, '(I5)') n
    print *, n
end program test
"#);
}

#[test] fn read_multiple_from_string() {
    compile_ok(r#"
program test
    character(len=20) :: buf = '1 2 3'
    integer :: a, b, c
    read(buf, *) a, b, c
    print *, a + b + c
end program test
"#);
}

#[test] fn read_string_from_string() {
    compile_ok(r#"
program test
    character(len=20) :: buf = 'hello world'
    character(len=5) :: word
    read(buf, '(A5)') word
    print *, word
end program test
"#);
}

#[test] fn read_real_formatted() {
    compile_ok(r#"
program test
    character(len=15) :: buf = ' 3.14159'
    real :: x
    read(buf, '(F8.5)') x
    print *, x
end program test
"#);
}

// ── Write then read roundtrip ─────────────────────────────────

#[test] fn write_read_roundtrip_int() {
    compile_ok(r#"
program test
    character(len=20) :: buf
    integer :: x = 12345, y
    write(buf, '(I10)') x
    read(buf, '(I10)') y
    print *, x == y
end program test
"#);
}

#[test] fn write_read_roundtrip_real() {
    compile_ok(r#"
program test
    character(len=20) :: buf
    real :: x = 3.14, y
    write(buf, '(F10.4)') x
    read(buf, '(F10.4)') y
    print *, abs(x - y) < 1e-3
end program test
"#);
}

#[test] fn write_read_roundtrip_string() {
    compile_ok(r#"
program test
    character(len=20) :: buf
    character(len=5) :: s1 = 'world', s2
    write(buf, '(A5)') s1
    read(buf, '(A5)') s2
    print *, s1 == s2
end program test
"#);
}

// ── Internal I/O for number formatting ────────────────────────

#[test] fn format_integer_as_string() {
    compile_ok(r#"
program test
    character(len=10) :: s
    integer :: n = 255
    write(s, '(I0)') n
    print *, trim(s)
end program test
"#);
}

#[test] fn format_hex_as_string() {
    compile_ok(r#"
program test
    character(len=10) :: s
    write(s, '(Z8)') 255
    print *, trim(s)
end program test
"#);
}

#[test] fn format_scientific_as_string() {
    compile_ok(r#"
program test
    character(len=20) :: s
    write(s, '(E12.4)') 1.23456e10
    print *, trim(s)
end program test
"#);
}

#[test] fn build_csv_line() {
    compile_ok(r#"
program test
    character(len=60) :: line
    integer :: a = 1, b = 2, c = 3
    write(line, '(I0, A, I0, A, I0)') a, ',', b, ',', c
    print *, trim(line)
end program test
"#);
}

#[test] fn internal_io_in_loop() {
    compile_ok(r#"
program test
    character(len=10) :: bufs(5)
    integer :: i
    do i = 1, 5
        write(bufs(i), '(I0)') i * i
    end do
    do i = 1, 5
        print *, trim(bufs(i))
    end do
end program test
"#);
}

#[test] fn iostat_on_internal_read() {
    compile_ok(r#"
program test
    character(len=5) :: buf = 'abc'
    integer :: n, ios
    read(buf, *, iostat=ios) n
    if (ios /= 0) then
        print *, 'parse error'
    end if
end program test
"#);
}

#[test] fn internal_io_in_function() {
    compile_ok(r#"
program test
    character(len=20) :: s
    s = int_to_str(42)
    print *, trim(s)
contains
    function int_to_str(n) result(s)
        integer, intent(in) :: n
        character(len=20) :: s
        write(s, '(I0)') n
    end function int_to_str
end program test
"#);
}
