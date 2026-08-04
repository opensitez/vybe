! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_number_returns_unit
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: num
open(52, status='scratch')
inquire(unit=52, number=num)
close(52)
if ((num) /= 52) then
    print *, "FAIL: want [52] got [", num, "]"
    stop 1
end if
end program t
