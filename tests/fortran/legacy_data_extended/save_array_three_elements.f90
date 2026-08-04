! vybe-test: fortran/legacy_data_extended/save_array_three_elements
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
call vec_once()
contains
subroutine vec_once()
integer, save :: vec(3) = (/5, 6, 7/)
if ((vec(2)) /= 6) then
    print *, "FAIL: want [6] got [", vec(2), "]"
    stop 1
end if
if ((sum(vec)) /= 18) then
    print *, "FAIL: want [18] got [", sum(vec), "]"
    stop 1
end if
end subroutine vec_once
end program t
