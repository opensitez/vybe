! vybe-test: fortran/module_use_extended/compile_use_rename_and_only_combined
! origin: languages/fortran/tests/fortran/test_module_use_extended.rs

module symbols
    implicit none
    integer :: alpha = 1
    integer :: beta = 2
contains
    function combine() result(r)
        integer :: r
        r = alpha + beta
    end function combine
end module symbols

program t
    use symbols, only: total => combine
    print *, total()
end program t
