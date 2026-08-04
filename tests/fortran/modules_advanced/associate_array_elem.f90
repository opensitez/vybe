! vybe-test: fortran/modules_advanced/associate_array_elem
! origin: languages/fortran/tests/fortran/test_modules_advanced.rs

program test
    integer :: a(5) = [10, 20, 30, 40, 50]
    associate(mid => a(3))
        print *, mid
    end associate
end program test
