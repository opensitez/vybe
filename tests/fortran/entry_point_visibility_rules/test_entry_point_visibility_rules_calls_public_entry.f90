! vybe-test: fortran/entry_point_visibility_rules/test_entry_point_visibility_rules_calls_public_entry
! origin: languages/fortran/tests/fortran/test_entry_point_visibility_rules.rs

module entry_points
    public :: public_entry
    private :: private_entry

    contains

    integer function public_entry()
        public_entry = 9
    end function

    integer function private_entry()
        private_entry = 1
    end function
end module

program test_entry_point_visibility_rules
    use entry_points
    if ((public_entry()) /= 9) then
    print *, "FAIL: want [9] got [", public_entry(), "]"
    stop 1
end if
end program test_entry_point_visibility_rules
