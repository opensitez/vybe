! vybe-test: fortran/variable_declarations_extended/integer_selected_int_kind_decl
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: ik = selected_int_kind(9)
integer(kind=ik) :: v = 512
if ((v) /= 512) then
    print *, "FAIL: want [512] got [", v, "]"
    stop 1
end if
end program t
