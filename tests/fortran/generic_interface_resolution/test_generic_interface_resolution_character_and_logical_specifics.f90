! vybe-test: fortran/generic_interface_resolution/test_generic_interface_resolution_character_and_logical_specifics
! origin: languages/fortran/tests/fortran/test_generic_interface_resolution.rs

program test_generic_interface_resolution_character_and_logical_specifics
    if ((size_or_truth("abc")) /= 3) then
    print *, "FAIL: want [3] got [", size_or_truth("abc"), "]"
    stop 1
end if
    if ((size_or_truth(.true.)) /= 1) then
    print *, "FAIL: want [1] got [", size_or_truth(.true.), "]"
    stop 1
end if

contains
    interface size_or_truth
        module procedure size_text
        module procedure truth_int
    end interface

    integer function size_text(value)
        character(len=*), intent(in) :: value
        size_text = len_trim(value)
    end function

    integer function truth_int(value)
        logical, intent(in) :: value
        if (value) then
            truth_int = 1
        else
            truth_int = 0
        end if
    end function
end program test_generic_interface_resolution_character_and_logical_specifics
