! vybe-test: fortran/program_units/program_defined_io_13
! origin: languages/fortran/tests/fortran/test_program_units.rs
module m
type :: t
 integer :: x
contains
 procedure :: write_formatted
 generic :: write(formatted) => write_formatted
end type t
contains
subroutine write_formatted(dtv, unit, iotype, v_list, iostat, iomsg)
 class(t), intent(in) :: dtv
 integer, intent(in) :: unit
 character(len=*), intent(in) :: iotype
 integer, intent(in) :: v_list(:)
 integer, intent(out) :: iostat
 character(len=*), intent(inout) :: iomsg
 write(unit, '(i0)', iostat=iostat) dtv%x
end subroutine write_formatted
end module m
program driver
use m
type(t) :: obj
character(len=16) :: buf
obj%x = 41
write(buf, '(DT)') obj
if (trim(buf) /= "41") then
    print *, "FAIL: want [41] got [", trim(buf), "]"
    stop 1
end if
end program driver
