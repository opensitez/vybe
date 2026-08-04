! vybe-test: fortran/associate_construct_extended/associate_array_whole_vector
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(4) = [2, 4, 6, 8]
associate (vec => a)
if ((sum(vec)) /= 20) then
    print *, "FAIL: want [20] got [", sum(vec), "]"
    stop 1
end if
end associate
end program t
