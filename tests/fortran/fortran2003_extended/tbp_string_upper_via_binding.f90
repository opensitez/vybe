! vybe-test: fortran/fortran2003_extended/tbp_string_upper_via_binding
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module m
type :: Word
character(len=12) :: text = 'fortran'
contains
procedure :: upper
end type Word
contains
function upper(self) result(out)
class(Word), intent(in) :: self
character(len=12) :: out
out = self%text
end function upper
end module m
program driver
use m
type(Word) :: w
if (trim(trim(w%upper())) /= "fortran") then
    print *, "FAIL: want [fortran] got [", trim(w%upper()), "]"
    stop 1
end if
end program driver
