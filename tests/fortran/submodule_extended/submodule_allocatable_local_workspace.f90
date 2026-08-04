! vybe-test: fortran/submodule_extended/submodule_allocatable_local_workspace
! origin: languages/fortran/tests/fortran/test_submodule_extended.rs

module grow_iface
    implicit none
    interface
        module function grown(n) result(r)
            integer, intent(in) :: n
            integer :: r
        end function grown
    end interface
end module grow_iface

submodule (grow_iface) grow_impl
contains
    module function grown(n) result(r)
        integer, intent(in) :: n
        integer, allocatable :: buf(:)
        integer :: r
        allocate(buf(n))
        buf = 1
        r = sum(buf)
        deallocate(buf)
    end function grown
end submodule grow_impl

program t
    use grow_iface
    if ((grown(5)) /= 5) then
    print *, "FAIL: want [5] got [", grown(5), "]"
    stop 1
end if
end program t
