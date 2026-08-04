! vybe-test: fortran/submodule_extended/submodule_intent_out_fills_vector
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 4 ]

module fill_iface
    implicit none
    interface
        module subroutine fill_seq(v, n)
            integer, intent(out) :: v(:)
            integer, intent(in) :: n
        end subroutine fill_seq
    end interface
end module fill_iface

submodule (fill_iface) fill_impl
contains
    module subroutine fill_seq(v, n)
        integer, intent(out) :: v(:)
        integer, intent(in) :: n
        integer :: i
        do i = 1, n
            v(i) = i * 2
        end do
    end subroutine fill_seq
end submodule fill_impl

program t
    use fill_iface
    integer :: a(3)
    call fill_seq(a, 3)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((a(2)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", a(2), "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
