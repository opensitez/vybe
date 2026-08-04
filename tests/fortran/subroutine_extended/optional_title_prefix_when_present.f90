! vybe-test: fortran/subroutine_extended/optional_title_prefix_when_present
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
call greet_opt('Ann')
call greet_opt('Ann', 'Ms.')
contains
subroutine greet_opt(name, title)
character(len=*), intent(in) :: name
character(len=*), intent(in), optional :: title
if (present(title)) then
if (trim(trim(title) // ' ' // trim(name)) /= "Ann") then
    print *, "FAIL: want [Ann] got [", trim(title) // ' ' // trim(name), "]"
    stop 1
end if
else
if (trim(trim(name)) /= "Ms. Ann") then
    print *, "FAIL: want [Ms. Ann] got [", trim(name), "]"
    stop 1
end if
end if
end subroutine greet_opt
end program t
