! vybe-test: fortran/module_use_extended/compile_use_only_triple_symbol_list
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module trio
    implicit none
    integer :: u = 1, v = 2, w = 3
contains
    function sum3() result(r)
        integer :: r
        r = u + v + w
    end function sum3
end module trio

program t
    use trio, only: u, w, sum3
    print *, sum3()
end program t
