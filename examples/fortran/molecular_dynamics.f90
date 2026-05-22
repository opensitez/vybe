! ============================================================
! Molecular Dynamics Simulation — Lennard-Jones Fluid
! ============================================================
! A complete MD simulation of a Lennard-Jones fluid in a
! periodic box. Computes thermodynamic properties, radial
! distribution function, and mean-square displacement.
!
! This is the kind of code that runs on supercomputers for
! materials science, drug discovery, and protein folding.
!
! Features:
!   - Verlet neighbor list for O(N) force calculation
!   - Velocity Verlet integrator
!   - Nose-Hoover thermostat
!   - Periodic boundary conditions with minimum image
!   - Radial distribution function g(r)
!   - Mean-square displacement (diffusion coefficient)
!   - Energy/pressure/temperature diagnostics
!   - XYZ trajectory file output
! ============================================================

module md_constants
    implicit none
    integer, parameter :: dp = kind(1.0d0)
    real(dp), parameter :: KB = 1.380649e-23_dp   ! Boltzmann constant (J/K)
    real(dp), parameter :: PI = 4.0_dp * atan(1.0_dp)
end module md_constants


module md_types
    use md_constants
    implicit none

    type :: particle
        real(dp) :: r(3)    ! position
        real(dp) :: v(3)    ! velocity
        real(dp) :: f(3)    ! force
        real(dp) :: r0(3)   ! unwrapped position (for MSD)
        real(dp) :: mass
    end type particle

    type :: neighbor_list
        integer, allocatable :: list(:,:)   ! list(i, 1:nneigh(i))
        integer, allocatable :: nneigh(:)   ! number of neighbors of i
        integer :: max_neigh
        real(dp) :: r_cut_sq, r_skin_sq
    end type neighbor_list

    type :: md_system
        type(particle), allocatable :: atoms(:)
        integer  :: n                    ! number of atoms
        real(dp) :: box(3)               ! box lengths
        real(dp) :: rho                  ! number density
        real(dp) :: T_target             ! target temperature
        real(dp) :: T_current            ! current temperature
        real(dp) :: P_current            ! current pressure
        real(dp) :: E_kin                ! kinetic energy
        real(dp) :: E_pot                ! potential energy
        real(dp) :: E_tot                ! total energy
        real(dp) :: virial               ! virial for pressure
        real(dp) :: dt                   ! timestep
        integer  :: step                 ! current step
        ! Nose-Hoover thermostat
        real(dp) :: xi                   ! thermostat variable
        real(dp) :: Q                    ! thermostat mass
        ! LJ parameters (reduced units: eps=1, sigma=1)
        real(dp) :: eps, sigma, r_cut
        type(neighbor_list) :: nlist
    end type md_system

end module md_types


module md_core
    use md_constants
    use md_types
    implicit none

contains

    ! Apply periodic boundary conditions (minimum image convention)
    pure subroutine pbc(dr, box)
        real(dp), intent(inout) :: dr(3)
        real(dp), intent(in)    :: box(3)
        integer :: k
        do k = 1, 3
            dr(k) = dr(k) - box(k) * nint(dr(k) / box(k))
        end do
    end subroutine pbc

    ! Wrap position into box
    pure subroutine wrap_position(r, box)
        real(dp), intent(inout) :: r(3)
        real(dp), intent(in)    :: box(3)
        integer :: k
        do k = 1, 3
            r(k) = r(k) - box(k) * floor(r(k) / box(k))
        end do
    end subroutine wrap_position

    ! Initialize FCC lattice
    subroutine init_fcc_lattice(sys)
        type(md_system), intent(inout) :: sys
        integer :: n_cell, i, j, k, idx
        real(dp) :: a, basis(3,4)
        real(dp) :: r(3)

        ! FCC basis vectors (4 atoms per unit cell)
        basis(:,1) = [0.0_dp, 0.0_dp, 0.0_dp]
        basis(:,2) = [0.5_dp, 0.5_dp, 0.0_dp]
        basis(:,3) = [0.5_dp, 0.0_dp, 0.5_dp]
        basis(:,4) = [0.0_dp, 0.5_dp, 0.5_dp]

        n_cell = nint((sys%n / 4.0_dp)**(1.0_dp/3.0_dp))
        a = sys%box(1) / n_cell

        idx = 0
        outer: do k = 0, n_cell - 1
            do j = 0, n_cell - 1
                do i = 0, n_cell - 1
                    do concurrent (m = 1:4)
                        if (idx + m <= sys%n) then
                            sys%atoms(idx+m)%r(1) = (i + basis(1,m)) * a
                            sys%atoms(idx+m)%r(2) = (j + basis(2,m)) * a
                            sys%atoms(idx+m)%r(3) = (k + basis(3,m)) * a
                            sys%atoms(idx+m)%r0   = sys%atoms(idx+m)%r
                        end if
                    end do
                    idx = idx + 4
                    if (idx >= sys%n) exit outer
                end do
            end do
        end do outer
    end subroutine init_fcc_lattice

    ! Maxwell-Boltzmann velocity initialization
    subroutine init_velocities(sys, seed)
        type(md_system), intent(inout) :: sys
        integer, intent(in) :: seed
        integer :: i, s
        real(dp) :: vcm(3), sigma_v, u1, u2

        sigma_v = sqrt(sys%T_target)
        s = seed
        vcm = 0.0_dp

        do i = 1, sys%n
            ! Box-Muller for each component
            s = mod(s * 1664525 + 1013904223, 2**30)
            u1 = real(s, dp) / 2.0_dp**30 + 1.0e-15_dp
            s = mod(s * 1664525 + 1013904223, 2**30)
            u2 = real(s, dp) / 2.0_dp**30
            sys%atoms(i)%v(1) = sigma_v * sqrt(-2.0_dp*log(u1)) * cos(2.0_dp*PI*u2)

            s = mod(s * 1664525 + 1013904223, 2**30)
            u1 = real(s, dp) / 2.0_dp**30 + 1.0e-15_dp
            s = mod(s * 1664525 + 1013904223, 2**30)
            u2 = real(s, dp) / 2.0_dp**30
            sys%atoms(i)%v(2) = sigma_v * sqrt(-2.0_dp*log(u1)) * cos(2.0_dp*PI*u2)

            s = mod(s * 1664525 + 1013904223, 2**30)
            u1 = real(s, dp) / 2.0_dp**30 + 1.0e-15_dp
            s = mod(s * 1664525 + 1013904223, 2**30)
            u2 = real(s, dp) / 2.0_dp**30
            sys%atoms(i)%v(3) = sigma_v * sqrt(-2.0_dp*log(u1)) * cos(2.0_dp*PI*u2)

            vcm = vcm + sys%atoms(i)%v
        end do

        ! Remove center-of-mass drift
        vcm = vcm / sys%n
        do i = 1, sys%n
            sys%atoms(i)%v = sys%atoms(i)%v - vcm
        end do
    end subroutine init_velocities

    ! Build Verlet neighbor list
    subroutine build_neighbor_list(sys)
        type(md_system), intent(inout) :: sys
        integer :: i, j
        real(dp) :: dr(3), r2

        sys%nlist%nneigh = 0

        do i = 1, sys%n - 1
            do j = i + 1, sys%n
                dr = sys%atoms(j)%r - sys%atoms(i)%r
                call pbc(dr, sys%box)
                r2 = sum(dr**2)
                if (r2 < sys%nlist%r_skin_sq) then
                    sys%nlist%nneigh(i) = sys%nlist%nneigh(i) + 1
                    if (sys%nlist%nneigh(i) <= sys%nlist%max_neigh) then
                        sys%nlist%list(i, sys%nlist%nneigh(i)) = j
                    end if
                    sys%nlist%nneigh(j) = sys%nlist%nneigh(j) + 1
                    if (sys%nlist%nneigh(j) <= sys%nlist%max_neigh) then
                        sys%nlist%list(j, sys%nlist%nneigh(j)) = i
                    end if
                end if
            end do
        end do
    end subroutine build_neighbor_list

    ! Compute LJ forces and potential energy
    subroutine compute_forces(sys)
        type(md_system), intent(inout) :: sys
        integer  :: i, j, k, nn
        real(dp) :: dr(3), r2, r2i, r6i, fij, phi
        real(dp) :: r_cut2, eps4, eps24
        real(dp) :: phi_cut   ! LJ potential at cutoff (for energy shift)

        r_cut2 = sys%r_cut**2
        eps4   = 4.0_dp * sys%eps
        eps24  = 24.0_dp * sys%eps

        ! LJ at cutoff for energy shift
        r6i      = (sys%sigma**2 / r_cut2)**3
        phi_cut  = eps4 * (r6i**2 - r6i)

        ! Zero forces
        do i = 1, sys%n
            sys%atoms(i)%f = 0.0_dp
        end do

        sys%E_pot  = 0.0_dp
        sys%virial = 0.0_dp

        do i = 1, sys%n
            nn = sys%nlist%nneigh(i)
            do k = 1, min(nn, sys%nlist%max_neigh)
                j = sys%nlist%list(i, k)
                if (j <= i) cycle   ! avoid double counting

                dr = sys%atoms(j)%r - sys%atoms(i)%r
                call pbc(dr, sys%box)
                r2 = sum(dr**2)

                if (r2 < r_cut2) then
                    r2i = sys%sigma**2 / r2
                    r6i = r2i**3
                    phi = eps4 * (r6i**2 - r6i) - phi_cut
                    fij = eps24 * r2i * (2.0_dp*r6i**2 - r6i)

                    sys%atoms(i)%f = sys%atoms(i)%f + fij * dr
                    sys%atoms(j)%f = sys%atoms(j)%f - fij * dr

                    sys%E_pot  = sys%E_pot  + phi
                    sys%virial = sys%virial + fij * r2
                end if
            end do
        end do
    end subroutine compute_forces

    ! Velocity Verlet integrator with Nose-Hoover thermostat
    subroutine velocity_verlet_nh(sys)
        type(md_system), intent(inout) :: sys
        integer  :: i
        real(dp) :: dt, dt2, alpha, T_inst, dof

        dt  = sys%dt
        dt2 = dt * 0.5_dp
        dof = 3.0_dp * (sys%n - 1)   ! degrees of freedom

        ! Half-step velocity update (with thermostat)
        T_inst = 2.0_dp * sys%E_kin / dof
        alpha  = 1.0_dp + dt2 * sys%xi

        do i = 1, sys%n
            sys%atoms(i)%v = (sys%atoms(i)%v + dt2 * sys%atoms(i)%f / sys%atoms(i)%mass) / alpha
        end do

        ! Full-step position update
        do i = 1, sys%n
            sys%atoms(i)%r  = sys%atoms(i)%r  + dt * sys%atoms(i)%v
            sys%atoms(i)%r0 = sys%atoms(i)%r0 + dt * sys%atoms(i)%v  ! unwrapped
            call wrap_position(sys%atoms(i)%r, sys%box)
        end do

        ! Recompute forces
        call compute_forces(sys)

        ! Compute kinetic energy
        sys%E_kin = 0.0_dp
        do i = 1, sys%n
            sys%E_kin = sys%E_kin + 0.5_dp * sys%atoms(i)%mass * sum(sys%atoms(i)%v**2)
        end do

        ! Second half-step velocity update
        T_inst = 2.0_dp * sys%E_kin / dof
        sys%xi = sys%xi + dt2 * (T_inst / sys%T_target - 1.0_dp) / sys%Q

        alpha = 1.0_dp + dt2 * sys%xi
        do i = 1, sys%n
            sys%atoms(i)%v = (sys%atoms(i)%v + dt2 * sys%atoms(i)%f / sys%atoms(i)%mass) / alpha
        end do

        ! Update kinetic energy after second half-step
        sys%E_kin = 0.0_dp
        do i = 1, sys%n
            sys%E_kin = sys%E_kin + 0.5_dp * sys%atoms(i)%mass * sum(sys%atoms(i)%v**2)
        end do

        sys%T_current = 2.0_dp * sys%E_kin / dof
        sys%E_tot     = sys%E_kin + sys%E_pot
        sys%P_current = sys%rho * sys%T_current + sys%virial / (3.0_dp * product(sys%box))

        sys%step = sys%step + 1
    end subroutine velocity_verlet_nh

    ! Radial distribution function
    subroutine compute_rdf(sys, r_max, n_bins, rdf, r_vals)
        type(md_system), intent(in)  :: sys
        real(dp),        intent(in)  :: r_max
        integer,         intent(in)  :: n_bins
        real(dp),        intent(out) :: rdf(n_bins), r_vals(n_bins)
        integer, allocatable :: hist(:)
        real(dp) :: dr_bin, r, r2, vol_shell, rho_ideal
        real(dp) :: dr(3)
        integer  :: i, j, bin

        allocate(hist(n_bins))
        hist   = 0
        dr_bin = r_max / n_bins

        do i = 1, sys%n - 1
            do j = i + 1, sys%n
                dr = sys%atoms(j)%r - sys%atoms(i)%r
                call pbc(dr, sys%box)
                r2 = sum(dr**2)
                r  = sqrt(r2)
                if (r < r_max) then
                    bin = int(r / dr_bin) + 1
                    if (bin <= n_bins) hist(bin) = hist(bin) + 2
                end if
            end do
        end do

        ! Normalize
        do bin = 1, n_bins
            r         = (bin - 0.5_dp) * dr_bin
            r_vals(bin) = r
            vol_shell = 4.0_dp * PI * r**2 * dr_bin
            rho_ideal = sys%rho * vol_shell * sys%n
            rdf(bin)  = hist(bin) / rho_ideal
        end do

        deallocate(hist)
    end subroutine compute_rdf

    ! Mean-square displacement
    real(dp) function compute_msd(sys)
        type(md_system), intent(in) :: sys
        integer :: i
        real(dp) :: dr(3)
        compute_msd = 0.0_dp
        do i = 1, sys%n
            dr = sys%atoms(i)%r0 - sys%atoms(i)%r0   ! placeholder — use stored r0
            ! In a real simulation, r0 is the initial position
            compute_msd = compute_msd + sum((sys%atoms(i)%r0 - sys%atoms(i)%r)**2)
        end do
        compute_msd = compute_msd / sys%n
    end function compute_msd

end module md_core


program molecular_dynamics
    use md_constants
    use md_types
    use md_core
    implicit none

    type(md_system) :: sys
    integer :: i, step, n_steps, n_equil, output_freq, rdf_freq
    integer :: traj_unit, diag_unit
    real(dp), allocatable :: rdf(:), r_vals(:), r0_init(:,:)
    integer, parameter :: N_RDF = 100
    real(dp) :: rdf_accum(N_RDF), r_vals_rdf(N_RDF)
    integer  :: n_rdf_samples
    real(dp) :: msd_sum, msd_count

    ! ── System parameters (reduced LJ units: eps=1, sigma=1, m=1) ────────────
    sys%n        = 256          ! number of atoms (must be 4*n_cell^3)
    sys%rho      = 0.8_dp       ! reduced density
    sys%T_target = 1.0_dp       ! reduced temperature
    sys%dt       = 0.005_dp     ! reduced time step
    sys%eps      = 1.0_dp
    sys%sigma    = 1.0_dp
    sys%r_cut    = 2.5_dp       ! standard LJ cutoff
    sys%xi       = 0.0_dp       ! thermostat variable
    sys%Q        = 10.0_dp      ! thermostat mass
    sys%step     = 0

    n_steps     = 5000
    n_equil     = 1000
    output_freq = 100
    rdf_freq    = 50

    ! Box length from density
    sys%box = (sys%n / sys%rho)**(1.0_dp/3.0_dp)

    ! Allocate atoms
    allocate(sys%atoms(sys%n))
    do i = 1, sys%n
        sys%atoms(i)%mass = 1.0_dp
        sys%atoms(i)%f    = 0.0_dp
        sys%atoms(i)%v    = 0.0_dp
    end do

    ! Neighbor list
    sys%nlist%max_neigh = 200
    sys%nlist%r_cut_sq  = sys%r_cut**2
    sys%nlist%r_skin_sq = (sys%r_cut + 0.5_dp)**2
    allocate(sys%nlist%list(sys%n, sys%nlist%max_neigh))
    allocate(sys%nlist%nneigh(sys%n))

    ! Initialize
    print *, "probe:init_fcc_lattice:start"
    call init_fcc_lattice(sys)
    print *, "probe:init_fcc_lattice:done"
    print *, "probe:atom1=", sys%atoms(1)%r(1), sys%atoms(1)%r(2), sys%atoms(1)%r(3)
    print *, "probe:atom2=", sys%atoms(2)%r(1), sys%atoms(2)%r(2), sys%atoms(2)%r(3)
    print *, "probe:init_velocities:start"
    call init_velocities(sys, seed=42)
    print *, "probe:init_velocities:done"
    print *, "probe:build_neighbor_list:start"
    call build_neighbor_list(sys)
    print *, "probe:build_neighbor_list:done"
    print *, "probe:compute_forces:start"
    call compute_forces(sys)
    print *, "probe:compute_forces:done"

    ! Compute initial KE
    print *, "probe:initial_ke:start"
    sys%E_kin = 0.0_dp
    do i = 1, sys%n
        sys%E_kin = sys%E_kin + 0.5_dp * sys%atoms(i)%mass * sum(sys%atoms(i)%v**2)
    end do
    print *, "probe:initial_ke:done"

    ! Store initial positions for MSD
    allocate(r0_init(3, sys%n))
    do i = 1, sys%n
        r0_init(:, i) = sys%atoms(i)%r
    end do

    rdf_accum    = 0.0_dp
    n_rdf_samples = 0
    msd_sum      = 0.0_dp
    msd_count    = 0.0_dp

    ! Open output files
    print *, "probe:open:start"
    open(newunit=traj_unit, file="trajectory.xyz", status="replace")
    open(newunit=diag_unit, file="diagnostics.dat", status="replace")
    print *, "probe:open:done"
    write(diag_unit, "(a)") "# step  time  T  P  E_kin  E_pot  E_tot  MSD"

    print *, "============================================"
    print *, " Lennard-Jones Molecular Dynamics"
    print *, "============================================"
    print "(a, i0)", "  N atoms    = ", sys%n
    print "(a, f8.4)", "  Density    = ", sys%rho
    print "(a, f8.4)", "  T_target   = ", sys%T_target
    print "(a, f8.4)", "  Box length = ", sys%box(1)
    print "(a, i0, a, i0)", "  Steps: ", n_equil, " equil + ", n_steps - n_equil, " production"
    print *, ""
    print "(a6, a8, 4a12, a10)", "Step", "Time", "T", "P", "E_kin", "E_pot", "MSD"
    print "(a6, a8, 4a12, a10)", "----", "----", "-", "-", "-----", "-----", "---"

    ! ── Main MD loop ──────────────────────────────────────────────────────────
    do step = 1, n_steps

        ! Rebuild neighbor list every 20 steps
        if (mod(step, 20) == 0) call build_neighbor_list(sys)

        call velocity_verlet_nh(sys)

        ! MSD (production phase only)
        if (step > n_equil) then
            block
                real(dp) :: msd_val, dr(3)
                integer  :: ii
                msd_val = 0.0_dp
                do ii = 1, sys%n
                    dr = sys%atoms(ii)%r0 - r0_init(:, ii)
                    msd_val = msd_val + sum(dr**2)
                end do
                msd_sum   = msd_sum + msd_val / sys%n
                msd_count = msd_count + 1.0_dp
            end block
        end if

        ! Accumulate RDF (production phase)
        if (step > n_equil .and. mod(step, rdf_freq) == 0) then
            call compute_rdf(sys, sys%box(1)*0.5_dp, N_RDF, rdf, r_vals_rdf)
            rdf_accum     = rdf_accum + rdf
            n_rdf_samples = n_rdf_samples + 1
        end if

        ! Diagnostics
        if (mod(step, output_freq) == 0) then
            block
                real(dp) :: msd_now, dr(3)
                integer  :: ii
                msd_now = 0.0_dp
                do ii = 1, sys%n
                    dr = sys%atoms(ii)%r0 - r0_init(:, ii)
                    msd_now = msd_now + sum(dr**2)
                end do
                msd_now = msd_now / sys%n

                print "(i6, f8.3, 4f12.4, f10.4)", &
                    step, step * sys%dt, sys%T_current, sys%P_current, &
                    sys%E_kin / sys%n, sys%E_pot / sys%n, msd_now

                write(diag_unit, "(i6, 7es14.6)") &
                    step, step * sys%dt, sys%T_current, sys%P_current, &
                    sys%E_kin / sys%n, sys%E_pot / sys%n, &
                    sys%E_tot / sys%n, msd_now
            end block
        end if

        ! Write trajectory (XYZ format)
        if (mod(step, output_freq*5) == 0) then
            write(traj_unit, "(i0)") sys%n
            write(traj_unit, "(a, i0, a, f8.3)") "step=", step, " t=", step*sys%dt
            do i = 1, sys%n
                write(traj_unit, "(a2, 3f12.6)") "Ar", sys%atoms(i)%r
            end do
        end if

    end do

    close(traj_unit)
    close(diag_unit)

    ! ── Output RDF ────────────────────────────────────────────────────────────
    if (n_rdf_samples > 0) then
        rdf_accum = rdf_accum / n_rdf_samples
        print *, ""
        print *, "=== Radial Distribution Function g(r) ==="
        print "(a8, a10)", "r/sigma", "g(r)"
        do i = 1, N_RDF
            if (r_vals_rdf(i) > 0.5_dp) then
                print "(f8.3, f10.4)", r_vals_rdf(i), rdf_accum(i)
            end if
        end do
    end if

    ! ── Diffusion coefficient from MSD ────────────────────────────────────────
    if (msd_count > 0) then
        print *, ""
        print "(a, es12.4)", "Mean MSD (production) = ", msd_sum / msd_count
        print "(a, es12.4)", "Diffusion coeff D ~ MSD/(6t) = ", &
            msd_sum / msd_count / (6.0_dp * (n_steps - n_equil) * sys%dt)
    end if

    print *, ""
    print *, "Output files: trajectory.xyz, diagnostics.dat"
    print *, "Simulation complete."

    deallocate(sys%atoms, sys%nlist%list, sys%nlist%nneigh, r0_init)
    if (allocated(rdf)) deallocate(rdf)

end program molecular_dynamics
