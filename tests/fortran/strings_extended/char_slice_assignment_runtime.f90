! vybe-test: fortran/strings_extended/char_slice_assignment_runtime
! origin: languages/fortran/tests/fortran/test_strings_extended.rs
program t
character(len=5) :: s = 'abcde'
s(2:4) = 'XYZ'
if (trim(trim(s)) /= "aXYZe") then
    print *, "FAIL: want [aXYZe] got [", trim(s), "]"
    stop 1
end if
end program t
