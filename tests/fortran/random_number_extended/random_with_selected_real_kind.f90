! vybe-test: fortran/random_number_extended/random_with_selected_real_kind
! origin: languages/fortran/tests/fortran/test_random_number_extended.rs
program t
integer, parameter :: sp = selected_real_kind(6)
real(sp) :: r
call random_number(r)
if ((merge(1, 0, r >= 0.0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, r >= 0.0), "]"
    stop 1
end if
end program t
