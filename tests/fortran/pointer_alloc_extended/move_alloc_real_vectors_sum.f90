! vybe-test: fortran/pointer_alloc_extended/move_alloc_real_vectors_sum
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
real, allocatable :: a(:), b(:)
allocate(a(4))
a = [1.0, 2.0, 3.0, 4.0]
call move_alloc(a, b)
if ((sum(b)) /= 10) then
    print *, "FAIL: want [10] got [", sum(b), "]"
    stop 1
end if
if ((allocated(a)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(a), "]"
    stop 1
end if
end program t
