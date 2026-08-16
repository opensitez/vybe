! vybe-test: fortran/kind_inquiry/kind_of_kind_result_is_four
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
if ((kind(kind(1))) /= 4) then
    print *, "FAIL: want [4] got [", kind(kind(1)), "]"
    stop 1
end if
end program t
