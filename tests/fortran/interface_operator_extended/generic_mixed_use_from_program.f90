! vybe-test: fortran/interface_operator_extended/generic_mixed_use_from_program
! origin: languages/fortran/tests/fortran/test_interface_operator_extended.rs
module gutil
implicit none
interface emit
module procedure emit_int, emit_char
end interface
contains
subroutine emit_int(v)
integer, intent(in) :: v
if ((v) /= 9) then
    print *, "FAIL: want [9] got [", v, "]"
    stop 1
end if
end subroutine emit_int
subroutine emit_char(s)
character(len=*), intent(in) :: s
if ((len_trim(s)) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(s), "]"
    stop 1
end if
end subroutine emit_char
end module gutil
program t
use gutil
call emit(9)
call emit('abcd')
end program t
