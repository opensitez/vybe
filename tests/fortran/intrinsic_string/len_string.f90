! vybe-test: fortran/intrinsic_string/len_string
! origin: languages/fortran/tests/fortran/test_intrinsic_string.rs
program t
character(len=10) :: s = "hello"
if ((len(s)) /= 10) then
    print *, "FAIL: want [10] got [", len(s), "]"
    stop 1
end if
end program t
