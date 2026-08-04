! vybe-test: fortran/interface_blocks/interface_blocks_18
! origin: languages/fortran/tests/fortran/test_interface_blocks.rs
module m
implicit none

interface add_one
module procedure add_one_impl
end interface
contains
function add_one_impl(x) result(r)
integer, intent(in) :: x
integer :: r
r = x + 1
end function add_one_impl
end module m
