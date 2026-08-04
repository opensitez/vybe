! vybe-test: fortran/where_advanced/nested_where_2d
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    real :: m(4,4) = reshape([(real(i-8), i=1,16)],[4,4])
    real :: result(4,4)
    result = 0.0
    where (m > 0.0)
        where (m > 4.0)
            result = m * 2.0
        elsewhere
            result = m
        end where
    end where
    print *, result(1,1)
end program test
