! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_type_is_exact_child_expected_miss
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 0 ]
    type :: Base
        integer :: a = 1
    end type Base
    type, extends(Base) :: Child
        integer :: b = 5
    end type Child

    class(Base), allocatable :: payload
    allocate(Child :: payload)
    select type (payload)
    type is (Base)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((payload%a) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", payload%a, "]"
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
end program t
