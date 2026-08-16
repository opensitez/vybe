! vybe-test: fortran/interfaces/if_defined_io_24
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type::t
integer::v=7
contains
procedure::wf
generic::write(formatted)=>wf
end type
contains
subroutine wf(dtv,unit,iotype,v_list,iostat,iomsg)
class(t),intent(in)::dtv
integer,intent(in)::unit
character(len=*),intent(in)::iotype
integer,intent(in)::v_list(:)
integer,intent(out)::iostat
character(len=*),intent(inout)::iomsg
write(unit,'(i0)',iostat=iostat) dtv%v
end
end module m
program driver
use m
type(t)::obj
character(len=16)::buf
write(buf,'(DT)') obj
if (trim(buf) /= "7") then
    print *, "FAIL: want [7] got [", trim(buf), "]"
    stop 1
end if
end program driver
