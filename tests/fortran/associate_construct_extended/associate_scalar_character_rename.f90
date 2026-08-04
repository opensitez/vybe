! vybe-test: fortran/associate_construct_extended/associate_scalar_character_rename
! origin: languages/fortran/tests/fortran/test_associate_construct_extended.rs
program t
character(len=6) :: word = 'fortran'
associate (w => word)
if (trim(trim(w)) /= "fortran") then
    print *, "FAIL: want [fortran] got [", trim(w), "]"
    stop 1
end if
end associate
end program t
