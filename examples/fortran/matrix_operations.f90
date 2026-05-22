! Matrix Operations — demonstrates arrays, do loops, functions, subroutines
! Covers: multi-dimensional arrays, array sections, intrinsics (matmul, transpose,
!         dot_product), formatted I/O, intent attributes, pure functions.

program matrix_operations
    implicit none

    integer, parameter :: N = 4
    real(8), dimension(N, N) :: A, B, C
    real(8), dimension(N)    :: v, w
    integer :: i, j

    ! Initialize matrix A with a non-symmetric pattern so transpose is obvious
    do i = 1, N
        do j = 1, N
            A(i, j) = real(10 * i + j, 8)
        end do
    end do

    ! Initialize matrix B as identity
    B = 0.0d0
    do i = 1, N
        B(i, i) = 1.0d0
    end do

    ! Matrix multiply using intrinsic
    C = matmul(A, B)

    print *, "=== Matrix A ==="
    call print_matrix(A, N)

    print *, ""
    print *, "=== A * I = A (should match above) ==="
    call print_matrix(C, N)

    print *, ""
    print *, "=== Transpose of A ==="
    call print_matrix(transpose(A), N)

    ! Vector operations using array constructor with implied do
    v = [(real(i, 8), i = 1, N)]
    w = matmul(A, v)

    print *, ""
    print *, "=== Vector v ==="
    print "(4f10.4)", v

    print *, "=== A * v ==="
    print "(4f10.4)", w

    print *, ""
    print "(a, f12.6)", "Dot product v.w    = ", dot_product(v, w)
    print "(a, f12.6)", "Frobenius norm A   = ", frobenius_norm(A, N)
    print "(a, f12.6)", "Trace of A         = ", matrix_trace(A, N)

    ! Array sections
    print *, ""
    print *, "=== First row of A ==="
    print "(4f10.4)", A(1, :)

    print *, "=== First column of A ==="
    print "(4f10.4)", A(:, 1)

    print *, "=== Sub-matrix A(2:3, 2:3) ==="
    call print_matrix(A(2:3, 2:3), 2)

contains

    subroutine print_matrix(M, n)
        integer, intent(in) :: n
        real(8), intent(in) :: M(n, n)
        integer :: i
        do i = 1, n
            print "(4f10.4)", M(i, :)
        end do
    end subroutine print_matrix

    pure function frobenius_norm(M, n) result(norm)
        integer, intent(in) :: n
        real(8), intent(in) :: M(n, n)
        real(8) :: norm
        norm = sqrt(sum(M**2))
    end function frobenius_norm

    pure function matrix_trace(M, n) result(tr)
        integer, intent(in) :: n
        real(8), intent(in) :: M(n, n)
        real(8) :: tr
        integer :: i
        tr = 0.0d0
        do i = 1, n
            tr = tr + M(i, i)
        end do
    end function matrix_trace

end program matrix_operations