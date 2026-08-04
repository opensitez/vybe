! vybe-test: fortran/generic_interfaces/test_generic_interfaces_character_concat_operator
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs

module m
    interface operator(//)
        module procedure mcat
    end interface
contains
    character(len=2) function mcat(a, b)
        character(len=*), intent(in) :: a, b
        mcat = a // b
    end function
end module m

program test_generic_interfaces_character_concat_operator
    use m
    if (trim('a' // 'b') /= "ab") then
    print *, "FAIL: want [ab] got [", 'a' // 'b', "]"
    stop 1
end if
end program test_generic_interfaces_character_concat_operator
