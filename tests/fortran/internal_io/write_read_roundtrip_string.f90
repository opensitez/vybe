! vybe-test: fortran/internal_io/write_read_roundtrip_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf
    character(len=5) :: s1 = 'world', s2
    write(buf, '(A5)') s1
    read(buf, '(A5)') s2
    print *, s1 == s2
end program test
