! vybe-test: fortran/select_case_extended/case_default_only
! origin: languages/fortran/tests/fortran/test_select_case_extended.rs
program t
integer :: n = 42
select case (n)
case default
if (trim("default") /= "default") then
    print *, "FAIL: want [default] got [", "default", "]"
    stop 1
end if
end select
end program t
