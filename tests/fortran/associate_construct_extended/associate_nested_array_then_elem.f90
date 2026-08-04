! vybe-test: fortran/associate_construct_extended/associate_nested_array_then_elem
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(4) = [2, 4, 6, 8]
associate (vec => a)
associate (elem => vec(3))
if ((elem) /= 6) then
    print *, "FAIL: want [6] got [", elem, "]"
    stop 1
end if
end associate
end associate
end program t
