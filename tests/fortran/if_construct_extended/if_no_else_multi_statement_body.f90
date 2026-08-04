! vybe-test: fortran/if_construct_extended/if_no_else_multi_statement_body
! origin: languages/fortran/tests/fortran/test_if_construct_extended.rs
program t
if (1 == 1) then
if (trim("step1") /= "step1") then
    print *, "FAIL: want [step1] got [", "step1", "]"
    stop 1
end if
if (trim("step2") /= "step2") then
    print *, "FAIL: want [step2] got [", "step2", "]"
    stop 1
end if
end if
end program t
