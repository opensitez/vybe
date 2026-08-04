! vybe-test: fortran/intrinsic_string/trim_string
! origin: languages/fortran/tests/fortran/test_intrinsic_string.rs
program t
if (trim(trim("  hello  ")) /= "hello") then
    print *, "FAIL: want [hello] got [", trim("  hello  "), "]"
    stop 1
end if
end program t
