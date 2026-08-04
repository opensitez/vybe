! vybe-test: fortran/substring_operations/substring_operations_runtime_variable_bounds
! origin: languages/fortran/tests/fortran/test_substring_operations.rs

program p
character(len=12) :: text
integer :: i
integer :: j
text = 'fortran-lang'
i = 2
j = 5
if (trim(trim(text(i:j))) /= "orta") then
    print *, "FAIL: want [orta] got [", trim(text(i:j)), "]"
    stop 1
end if
if ((len(text(i:j))) /= 4) then
    print *, "FAIL: want [4] got [", len(text(i:j)), "]"
    stop 1
end if
end program p
