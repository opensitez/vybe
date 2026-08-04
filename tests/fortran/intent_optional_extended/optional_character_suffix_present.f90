! vybe-test: fortran/intent_optional_extended/optional_character_suffix_present
! origin: languages/fortran/tests/fortran/test_intent_optional_extended.rs
program t
if (trim(trim(label('core', 'ext'))) /= "core_ext") then
    print *, "FAIL: want [core_ext] got [", trim(label('core', 'ext')), "]"
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
