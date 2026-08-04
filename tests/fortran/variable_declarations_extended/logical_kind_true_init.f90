! vybe-test: fortran/variable_declarations_extended/logical_kind_true_init
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical(kind=4) :: flag = .true.
if ((flag) .neqv. .true.) then
    print *, "FAIL: want [true] got [", flag, "]"
    stop 1
end if
end program t
