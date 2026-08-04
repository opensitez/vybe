! vybe-test: fortran/intrinsic_string/adjustl_string
! origin: languages/fortran/tests/fortran/test_intrinsic_string.rs
program t
if (trim(adjustl("  hello")) /= "hello") then
    print *, "FAIL: want [hello] got [", adjustl("  hello"), "]"
    stop 1
end if
end program t
