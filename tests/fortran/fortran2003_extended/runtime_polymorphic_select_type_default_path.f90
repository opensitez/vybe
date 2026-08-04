! vybe-test: fortran/fortran2003_extended/runtime_polymorphic_select_type_default_path
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 3 ]
    type :: Base
        integer :: n = 1
    end type Base
    type, extends(Base) :: Derived
        integer :: m = 3
    end type Derived
    type, extends(Base) :: Alt
        integer :: m = 5
    end type Alt

    class(Base), allocatable :: obj
    allocate(Derived :: obj)
    select type (obj)
    class is (Derived)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((obj%m) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", obj%m, "]"
            stop 1
        end if
    type is (Alt)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((-1) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", -1, "]"
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
