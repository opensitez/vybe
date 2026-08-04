! vybe-test: fortran/if_blocks/nested_if
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
if (1 > 0) then
if (2 > 1) then
if (trim("nested") /= "nested") then
    print *, "FAIL: want [nested] got [", "nested", "]"
    stop 1
end if
end if
end if
end program t
