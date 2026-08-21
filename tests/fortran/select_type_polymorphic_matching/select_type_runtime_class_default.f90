! vybe-test: fortran/select_type_polymorphic_matching/select_type_runtime_class_default
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 7 ]
    type :: Base
    end type Base
    type, extends(Base) :: Derived
        integer :: n = 7
    end type Derived

    class(Base), allocatable :: item
    allocate(Derived :: item)
    select type (item)
    type is (Derived)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((item%n) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", item%n, "]"
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
end program test
