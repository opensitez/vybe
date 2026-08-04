! vybe-test: fortran/if_blocks/if_ne
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (3 /= 4) then
if (trim("not equal") /= "not equal") then
    print *, "FAIL: want [not equal] got [", "not equal", "]"
    stop 1
end if
end if
end program t
