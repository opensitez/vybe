! vybe-test: fortran/intrinsic_string/len_trim_string
! origin: languages/fortran/tests/fortran/test_intrinsic_string.rs
program t
character(len=20) :: s = "hello"
if ((len_trim(s)) /= 5) then
    print *, "FAIL: want [5] got [", len_trim(s), "]"
    stop 1
end if
end program t
