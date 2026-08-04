! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_iostat_zero_on_success_open
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
integer :: ios
open(56, file='ioc_ext_ios.dat', status='replace', iostat=ios)
close(56, status='delete')
if ((ios) /= 0) then
    print *, "FAIL: want [0] got [", ios, "]"
    stop 1
end if
end program t
