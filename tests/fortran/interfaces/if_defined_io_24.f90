! vybe-test: fortran/interfaces/if_defined_io_24
! origin: languages/fortran/tests/fortran/test_interfaces.rs
module m
type::t
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
iostat=0
end
end module m
