! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_class_is_extension_chain
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_class_is_extension_chain
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
    type :: Base
        integer :: a = 1
    end type Base
    type, extends(Base) :: Child
        integer :: b = 5
    end type Child

    class(Base), allocatable :: item
    allocate(Child :: item)
    select type(item)
    class is (Child)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((item%b) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", item%b, "]"
            stop 1
        end if
    class is (Base)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((item%a) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", item%a, "]"
            stop 1
        end if
    class default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((-1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_type_polymorphic_matching_class_is_extension_chain
