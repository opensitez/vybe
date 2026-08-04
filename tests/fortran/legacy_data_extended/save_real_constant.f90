! vybe-test: fortran/legacy_data_extended/save_real_constant
! origin: languages/fortran/tests/fortran/test_legacy_data_extended.rs
program t
call pi_once()
contains
subroutine pi_once()
real, save :: pi = 3.25
if (abs((pi) - 3.25) > 1.0e-6) then
    print *, "FAIL: want [3.25] got [", pi, "]"
    stop 1
end if
end subroutine pi_once
end program t
