! vybe-test: fortran/associate_construct_extended/associate_index_from_expression
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a(10)
a = [(i, i = 1, 10)]
integer :: idx = 4
associate (elem => a(idx + 1))
if ((elem) /= 5) then
    print *, "FAIL: want [5] got [", elem, "]"
    stop 1
end if
end associate
end program t
