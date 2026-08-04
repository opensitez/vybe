! vybe-test: fortran/if_blocks/if_ge
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (5 >= 5) then
if (trim("yes") /= "yes") then
    print *, "FAIL: want [yes] got [", "yes", "]"
    stop 1
end if
end if
end program t
