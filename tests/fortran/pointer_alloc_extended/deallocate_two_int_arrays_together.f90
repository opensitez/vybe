! vybe-test: fortran/pointer_alloc_extended/deallocate_two_int_arrays_together
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: a(:), b(:)
allocate(a(2), b(3))
a = [10, 20]
b = [1, 2, 3]
if ((sum(a)) /= 30) then
    print *, "FAIL: want [30] got [", sum(a), "]"
    stop 1
end if
if ((sum(b)) /= 6) then
    print *, "FAIL: want [6] got [", sum(b), "]"
    stop 1
end if
deallocate(a, b)
if ((allocated(a)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(a), "]"
    stop 1
end if
if ((allocated(b)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(b), "]"
    stop 1
end if
end program t
