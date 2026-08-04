! vybe-test: fortran/parameterized_types/pdt_in_module
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

module pdt_mod
    implicit none
    type :: Tensor(rank, k)
        integer, len  :: rank
        integer, kind :: k
        real(k) :: components(rank)
    end type Tensor
contains
    subroutine zero(t)
        type(Tensor(*,*)), intent(inout) :: t
        t%components = 0
    end subroutine zero
end module pdt_mod

program test
    use pdt_mod
    type(Tensor(3,4)) :: v
    call zero(v)
    print *, v%components(1)
end program test
