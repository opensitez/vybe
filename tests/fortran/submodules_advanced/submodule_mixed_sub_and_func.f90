! vybe-test: fortran/submodules_advanced/submodule_mixed_sub_and_func
! origin: languages/fortran/tests/fortran/test_submodules_advanced.rs
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 32 ]

module io_iface
    implicit none
    interface
        module subroutine print_vec(v)
            real, intent(in) :: v(:)
        end subroutine print_vec
        module function dot(u, v) result(d)
            real, intent(in) :: u(:), v(:)
            real :: d
        end function dot
    end interface
end module io_iface

submodule (io_iface) io_impl
    implicit none
contains
    module subroutine print_vec(v)
        real, intent(in) :: v(:)
        integer :: i
        do i = 1, size(v)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((v(i)) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", v(i), "]"
                stop 1
            end if
        end do
    end subroutine print_vec

    module function dot(u, v) result(d)
        real, intent(in) :: u(:), v(:)
        real :: d
        d = sum(u * v)
    end function dot
end submodule io_impl

program test
    use io_iface
    real :: u(3) = [1.0, 2.0, 3.0]
    real :: v(3) = [4.0, 5.0, 6.0]
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((int(dot(u, v))) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", int(dot(u, v)), "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
