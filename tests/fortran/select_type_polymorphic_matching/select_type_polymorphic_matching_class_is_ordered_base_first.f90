! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_class_is_ordered_base_first
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_class_is_ordered_base_first
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 10 ]
    type :: Base
        integer :: a = 10
    end type Base
    type, extends(Base) :: Child
        integer :: b = 20
    end type Child

    class(Base), allocatable :: payload
    allocate(Child :: payload)
    select type (payload)
    class is (Base)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((payload%a) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", payload%a, "]"
            stop 1
        end if
    class is (Child)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((payload%b) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", payload%b, "]"
            stop 1
        end if
    class default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((0) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", 0, "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_type_polymorphic_matching_class_is_ordered_base_first
