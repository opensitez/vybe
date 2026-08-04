! vybe-test: fortran/types/interface_basic
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    interface
        function add(a, b) result(res)
            integer, intent(in) :: a, b
            integer :: res
        end function add
    end interface
    if (trim("ok") /= "ok") then
    print *, "FAIL: want [ok] got [", "ok", "]"
    stop 1
end if
end program test
