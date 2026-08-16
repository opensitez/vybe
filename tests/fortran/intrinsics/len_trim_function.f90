! vybe-test: fortran/intrinsics/len_trim_function
! origin: languages/fortran/tests/fortran/test_intrinsics.rs

program test
    character(len=20) :: s
    s = "hello"
    if ((len(s)) /= 20) then
    print *, "FAIL: want [20] got [", len(s), "]"
    stop 1
end if
end program test
