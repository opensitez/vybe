! vybe-test: fortran/stream_access_io/stream_module_helper_roundtrip
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
module iostream
implicit none
contains
subroutine write_pair(u, a, b)
integer, intent(in) :: u, a, b
write(u) a, b
end subroutine write_pair
subroutine read_pair(u, a, b)
integer, intent(in) :: u
integer, intent(out) :: a, b
read(u) a, b
end subroutine read_pair
end module iostream
program t
use iostream
integer :: x, y
open(14, status='scratch', access='stream', form='unformatted')
call write_pair(14, 8, 9)
rewind(14)
call read_pair(14, x, y)
close(14)
print *, x * y
end program t
