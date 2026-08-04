! vybe-test: fortran/variable_declarations_extended/implicit_none_logical_character
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
logical :: ok = .true.
character(len=2) :: ch = "ok"
if ((ok) .neqv. .true.) then
    print *, "FAIL: want [true] got [", ok, "]"
    stop 1
end if
if (trim(trim(ch)) /= "ok") then
    print *, "FAIL: want [ok] got [", trim(ch), "]"
    stop 1
end if
end program t
