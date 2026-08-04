! vybe-test: fortran/if_blocks/if_eq
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (3 == 3) then
if (trim("equal") /= "equal") then
    print *, "FAIL: want [equal] got [", "equal", "]"
    stop 1
end if
end if
end program t
