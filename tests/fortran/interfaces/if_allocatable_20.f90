! vybe-test: fortran/interfaces/if_allocatable_20
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer, allocatable :: a(:)
allocate(a(3))
a = 1
call s(a)
if (size(a) /= 5) then
    print *, "FAIL: want [5] got [", size(a), "]"
    stop 1
end if
if (sum(a) /= 10) then
    print *, "FAIL: want [10] got [", sum(a), "]"
    stop 1
end if
contains
subroutine s(x)
integer,allocatable::x(:)
deallocate(x)
allocate(x(5))
x = 2
end subroutine s
end program t
