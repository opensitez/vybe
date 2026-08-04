! vybe-test: fortran/if_blocks/if_after_assignment
! origin: languages/fortran/tests/fortran/test_if_blocks.rs
program t
integer :: x = 10
x = x + 5
if (x == 15) then
if (trim("correct") /= "correct") then
    print *, "FAIL: want [correct] got [", "correct", "]"
    stop 1
end if
end if
end program t
