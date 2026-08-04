! vybe-test: fortran/intent_optional_extended/optional_character_suffix_missing
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if (trim(trim(label('core'))) /= "core") then
    print *, "FAIL: want [core] got [", trim(label('core')), "]"
    stop 1
end if
contains
character(len=20) function label(base, suffix)
character(len=*), intent(in) :: base
character(len=*), intent(in), optional :: suffix
if (present(suffix)) then
label = trim(base) // '_' // trim(suffix)
else
label = trim(base)
end if
end function label
end program t
