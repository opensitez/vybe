! vybe-test: fortran/fortran2018_extended/implicit_none_type_external_in_module
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

module guarded
    implicit none (type, external)
contains
    function twice(n) result(r)
        integer, intent(in) :: n
        integer :: r
        r = n * 2
    end function twice
end module guarded

program t
    use guarded
    print *, twice(11)
end program t
