! vybe-test: fortran/specification_part/spec_bindc_named
! origin: languages/fortran/tests/fortran/test_specification_part.rs
module m
implicit none
integer :: hits = 0
contains
subroutine s() bind(c, name='c_entry')
hits = hits + 1
end subroutine s
end module m
program t
use m
implicit none
call s()
if (hits /= 1) then
    print *, "FAIL: want [1] got [", hits, "]"
    stop 1
end if
end program t
