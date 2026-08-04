! vybe-test: fortran/pointer_alloc_extended/move_alloc_char_array_first_word
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
character(len=3), allocatable :: a(:), b(:)
allocate(a(2))
a(1) = 'xyz'
a(2) = 'uvw'
call move_alloc(a, b)
if (trim(trim(b(1))) /= "xyz") then
    print *, "FAIL: want [xyz] got [", trim(b(1)), "]"
    stop 1
end if
if ((allocated(a)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(a), "]"
    stop 1
end if
end program t
