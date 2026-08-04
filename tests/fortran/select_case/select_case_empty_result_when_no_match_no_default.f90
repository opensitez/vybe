! vybe-test: fortran/select_case/select_case_empty_result_when_no_match_no_default
! origin: languages/fortran/tests/fortran/test_select_case.rs
program t
integer :: n = 99
select case (n)
case (1)
print *, "one"
case (2)
print *, "two"
end select
end program t
