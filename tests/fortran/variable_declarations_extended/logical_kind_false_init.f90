! vybe-test: fortran/variable_declarations_extended/logical_kind_false_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical(kind=4) :: flag = .false.
if ((flag) .neqv. .false.) then
    print *, "FAIL: want [false] got [", flag, "]"
    stop 1
end if
end program t
