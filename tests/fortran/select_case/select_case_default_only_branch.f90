! vybe-test: fortran/select_case/select_case_default_only_branch
! origin: languages/fortran/tests/fortran/test_select_case.rs
program t
integer :: n = 99
select case (n)
case default
if (trim("fallback") /= "fallback") then
    print *, "FAIL: want [fallback] got [", "fallback", "]"
    stop 1
end if
end select
end program t
