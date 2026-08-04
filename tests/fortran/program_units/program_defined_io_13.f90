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
 iostat = 0
end subroutine write_formatted
end module m
