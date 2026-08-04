! vybe-test: fortran/deferred_length_operators/test_deferred_length_operators_on_allocatable_character
! origin: languages/fortran/tests/fortran/test_deferred_length_operators.rs

program test_deferred_length_operators
    character(len=:), allocatable :: text
    allocate(character(len=9) :: text)
    text = 'fortify-1'
    if ((len(text)) /= 9) then
    print *, "FAIL: want [9] got [", len(text), "]"
    stop 1
end if
    if (trim(trim(text)) /= "fortify-1") then
    print *, "FAIL: want [fortify-1] got [", trim(text), "]"
    stop 1
end if
end program test_deferred_length_operators
