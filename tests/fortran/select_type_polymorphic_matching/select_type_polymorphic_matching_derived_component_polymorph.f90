! vybe-test: fortran/select_type_polymorphic_matching/select_type_polymorphic_matching_derived_component_polymorph
! origin: languages/fortran/tests/fortran/test_select_type_polymorphic_matching.rs

program select_type_polymorphic_matching_derived_component_polymorph
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
    type :: Packet
        integer :: n = 2
    end type Packet
    class(*) :: holder
    class(Packet), allocatable :: payload
    allocate(Packet :: payload)
    holder = payload%n
    select type (payload)
    type is (Packet)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if ((payload%n) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", payload%n, "]"
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
end program select_type_polymorphic_matching_derived_component_polymorph
