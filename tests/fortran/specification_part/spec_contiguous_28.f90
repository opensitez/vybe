! vybe-test: fortran/specification_part/spec_contiguous_28
! origin: languages/fortran/tests/fortran/test_specification_part.rs
program t
implicit none
integer :: buf(3)
buf = [1, 2, 3]
call s(buf)
if (sum(buf) /= 12) then
    print *, "FAIL: want [12] got [", sum(buf), "]"
    stop 1
end if
contains
subroutine s(a)
implicit none
integer, contiguous :: a(:)
a = a * 2
end subroutine s
end program t
