! vybe-test: fortran/inquire_open_close_extended/ioc_iostat_close_already_closed_unit
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: ios
open(64, status='scratch')
close(64)
close(64, iostat=ios)
if ((ios) /= 0) then
    print *, "FAIL: want [0] got [", ios, "]"
    stop 1
end if
end program t
