! vybe-test: fortran/associate_construct_extended/associate_comparison_expr
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
integer :: a = 10, b = 20
associate (less => a < b)
if ((less) .neqv. .true.) then
    print *, "FAIL: want [true] got [", less, "]"
    stop 1
end if
end associate
end program t
