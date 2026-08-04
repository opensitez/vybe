! vybe-test: fortran/intrinsics/trim_function
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    if (trim(trim("  hello  ")) /= "hello") then
    print *, "FAIL: want [hello] got [", trim("  hello  "), "]"
    stop 1
end if
end program test
