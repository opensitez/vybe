! vybe-test: fortran/if_construct_extended/if_no_else_guard_positive_value
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
integer :: n = 5
if (n > 0) then
if (trim("ok") /= "ok") then
    print *, "FAIL: want [ok] got [", "ok", "]"
    stop 1
end if
end if
end program t
