! vybe-test: fortran/associate_construct_extended/associate_in_function_result
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
if ((double_it(6)) /= 12) then
    print *, "FAIL: want [12] got [", double_it(6), "]"
    stop 1
end if
contains
integer function double_it(n) result(r)
integer, intent(in) :: n
associate (twice => n * 2)
r = twice
end associate
end function double_it
end program t
