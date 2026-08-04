! vybe-test: fortran/parameterized_types/pdt_array_of_pdt
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Pair(k)
        integer, kind :: k
        real(k) :: x, y
    end type Pair
    type(Pair(4)) :: pairs(3)
    integer :: i
    do i = 1, 3
        pairs(i)%x = real(i, 4)
        pairs(i)%y = real(i, 4) * 2.0_4
    end do
    print *, pairs(2)%x
end program test
