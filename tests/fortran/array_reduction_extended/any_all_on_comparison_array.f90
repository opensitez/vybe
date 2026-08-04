! vybe-test: fortran/array_reduction_extended/any_all_on_comparison_array
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
integer :: a(5) = [2, 4, 6, 8, 10]
if ((any(a > 7)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", any(a > 7), "]"
    stop 1
end if
if ((all(a > 0)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", all(a > 0), "]"
    stop 1
end if
end program t
