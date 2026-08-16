! vybe-test: fortran/interfaces/if_contiguous_17
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
real :: buf(4)
buf = [1.0, 2.0, 3.0, 4.0]
call s(buf)
if (abs(sum(buf) - 20.0) > 1.0e-6) then
    print *, "FAIL: want [20.0] got [", sum(buf), "]"
    stop 1
end if
contains
subroutine s(a)
real,contiguous::a(:)
a = a * 2.0
end subroutine s
end program t
