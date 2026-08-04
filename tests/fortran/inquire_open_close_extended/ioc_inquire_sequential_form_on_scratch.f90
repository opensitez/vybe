! vybe-test: fortran/inquire_open_close_extended/ioc_inquire_sequential_form_on_scratch
! origin: languages/fortran/tests/fortran/test_inquire_open_close_extended.rs
program t
character(len=20) :: acc, frm
open(54, status='scratch')
inquire(unit=54, access=acc, form=frm)
close(54)
if (trim(acc(1:10)) /= "SEQUENTIAL") then
    print *, "FAIL: want [SEQUENTIAL] got [", acc(1:10), "]"
    stop 1
end if
end program t
