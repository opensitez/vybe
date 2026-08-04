! vybe-test: fortran/attributes/attr_protected_module_accessor
! origin: languages/fortran/tests/fortran/test_attributes.rs

module protected_access_mod
    integer, protected :: value = 9
contains
    integer function get_value()
        get_value = value
    end function get_value
end module protected_access_mod

program attr_protected_module_accessor
    use protected_access_mod
    if ((get_value()) /= 9) then
    print *, "FAIL: want [9] got [", get_value(), "]"
    stop 1
end if
end program attr_protected_module_accessor
