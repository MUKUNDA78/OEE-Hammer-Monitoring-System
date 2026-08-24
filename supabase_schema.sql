-- ==============================================================================
-- GRS ENGINEERING PVT LTD UNIT 1 - SUPABASE DATABASE SCHEMA & RLS POLICIES
-- Central Cloud PostgreSQL Database for 5-Hammer OEE & Quality Monitoring
-- ==============================================================================

-- Enable UUID Extension if not already enabled
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ------------------------------------------------------------------------------
-- 1. PROFILES TABLE (User Roles, Employee Names & Departments)
-- ------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.profiles (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    user_id UUID UNIQUE REFERENCES auth.users(id) ON DELETE CASCADE,
    employee_name TEXT NOT NULL,
    department TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('Admin', 'Quality', 'Production', 'Maintenance', 'Management', 'Viewer')),
    active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Index for profiles lookup
CREATE INDEX IF NOT EXISTS idx_profiles_user_id ON public.profiles(user_id);
CREATE INDEX IF NOT EXISTS idx_profiles_role ON public.profiles(role);

-- ------------------------------------------------------------------------------
-- 2. PRODUCTION DATA TABLE
-- ------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.production_data (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    date DATE NOT NULL,
    shift TEXT NOT NULL CHECK (shift IN ('Shift A', 'Shift B')),
    hammer TEXT NOT NULL CHECK (hammer IN ('1 Ton Hammer', '1.5 Ton Hammer', '2.5 Ton (Old) Hammer', '2.5 Ton (New) Hammer', '3.5 Ton Hammer')),
    part_number TEXT NOT NULL,
    planned_time_mins NUMERIC(10, 2) DEFAULT 660.00,
    planned_qty INT DEFAULT 0,
    production_qty INT NOT NULL DEFAULT 0,
    good_qty INT NOT NULL DEFAULT 0,
    ideal_cycle_sec NUMERIC(10, 2) DEFAULT 45.00,
    created_by UUID REFERENCES auth.users(id),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_by UUID REFERENCES auth.users(id),
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    is_deleted BOOLEAN DEFAULT FALSE,
    deleted_by UUID REFERENCES auth.users(id),
    deleted_at TIMESTAMPTZ
);

-- Indexes for production data filtering
CREATE INDEX IF NOT EXISTS idx_prod_date ON public.production_data(date);
CREATE INDEX IF NOT EXISTS idx_prod_shift ON public.production_data(shift);
CREATE INDEX IF NOT EXISTS idx_prod_hammer ON public.production_data(hammer);
CREATE INDEX IF NOT EXISTS idx_prod_part ON public.production_data(part_number);

-- ------------------------------------------------------------------------------
-- 3. QUALITY DATA TABLE (Maintains Rework & Rejection Separately)
-- ------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.quality_data (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    date DATE NOT NULL,
    shift TEXT NOT NULL CHECK (shift IN ('Shift A', 'Shift B')),
    hammer TEXT CHECK (hammer IN ('1 Ton Hammer', '1.5 Ton Hammer', '2.5 Ton (Old) Hammer', '2.5 Ton (New) Hammer', '3.5 Ton Hammer')),
    part_number TEXT NOT NULL,
    inspection_stage TEXT NOT NULL CHECK (inspection_stage IN ('In-Process', 'Final Inspection', 'MPI')),
    inspection_qty INT NOT NULL DEFAULT 0,
    rework_qty INT NOT NULL DEFAULT 0,
    rejection_qty INT NOT NULL DEFAULT 0,
    reason TEXT,
    rework_reason TEXT,
    rejection_reason TEXT,
    remarks TEXT,
    created_by UUID REFERENCES auth.users(id),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_by UUID REFERENCES auth.users(id),
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    is_deleted BOOLEAN DEFAULT FALSE,
    deleted_by UUID REFERENCES auth.users(id),
    deleted_at TIMESTAMPTZ
);

-- Indexes for quality data filtering
CREATE INDEX IF NOT EXISTS idx_qual_date ON public.quality_data(date);
CREATE INDEX IF NOT EXISTS idx_qual_stage ON public.quality_data(inspection_stage);
CREATE INDEX IF NOT EXISTS idx_qual_part ON public.quality_data(part_number);
CREATE INDEX IF NOT EXISTS idx_qual_reason ON public.quality_data(reason);

-- ------------------------------------------------------------------------------
-- 4. DOWNTIME DATA TABLE
-- ------------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS public.downtime_data (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    date DATE NOT NULL,
    shift TEXT NOT NULL CHECK (shift IN ('Shift A', 'Shift B')),
    hammer TEXT NOT NULL CHECK (hammer IN ('1 Ton Hammer', '1.5 Ton Hammer', '2.5 Ton (Old) Hammer', '2.5 Ton (New) Hammer', '3.5 Ton Hammer')),
    part_number TEXT NOT NULL,
    downtime_category TEXT NOT NULL CHECK (downtime_category IN ('Maintenance', 'Die Related', 'Setup', 'No Manpower', 'Minor Stop', 'Heating Time', 'Other')),
    downtime_minutes NUMERIC(10, 2) NOT NULL DEFAULT 0.00,
    reason TEXT,
    remarks TEXT,
    created_by UUID REFERENCES auth.users(id),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_by UUID REFERENCES auth.users(id),
    updated_at TIMESTAMPTZ DEFAULT NOW(),
    is_deleted BOOLEAN DEFAULT FALSE,
    deleted_by UUID REFERENCES auth.users(id),
    deleted_at TIMESTAMPTZ
);

-- Indexes for downtime data filtering
CREATE INDEX IF NOT EXISTS idx_down_date ON public.downtime_data(date);
CREATE INDEX IF NOT EXISTS idx_down_hammer ON public.downtime_data(hammer);
CREATE INDEX IF NOT EXISTS idx_down_cat ON public.downtime_data(downtime_category);

-- ------------------------------------------------------------------------------
-- 5. AUTOMATIC UPDATED_AT TRIGGER FUNCTION
-- ------------------------------------------------------------------------------
CREATE OR REPLACE FUNCTION public.handle_updated_at()
RETURNS TRIGGER AS $$
BEGIN
    NEW.updated_at = NOW();
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

-- Apply updated_at triggers
DROP TRIGGER IF EXISTS set_updated_at_profiles ON public.profiles;
CREATE TRIGGER set_updated_at_profiles BEFORE UPDATE ON public.profiles FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_production ON public.production_data;
CREATE TRIGGER set_updated_at_production BEFORE UPDATE ON public.production_data FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_quality ON public.quality_data;
CREATE TRIGGER set_updated_at_quality BEFORE UPDATE ON public.quality_data FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

DROP TRIGGER IF EXISTS set_updated_at_downtime ON public.downtime_data;
CREATE TRIGGER set_updated_at_downtime BEFORE UPDATE ON public.downtime_data FOR EACH ROW EXECUTE FUNCTION public.handle_updated_at();

-- ------------------------------------------------------------------------------
-- 6. AUTOMATIC USER PROFILE CREATION ON SIGNUP TRIGGER
-- ------------------------------------------------------------------------------
CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
    INSERT INTO public.profiles (user_id, employee_name, department, role)
    VALUES (
        NEW.id,
        COALESCE(NEW.raw_user_meta_data->>'employee_name', SPLIT_PART(NEW.email, '@', 1)),
        COALESCE(NEW.raw_user_meta_data->>'department', 'Production'),
        COALESCE(NEW.raw_user_meta_data->>'role', 'Production')
    )
    ON CONFLICT (user_id) DO NOTHING;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

DROP TRIGGER IF EXISTS on_auth_user_created ON auth.users;
CREATE TRIGGER on_auth_user_created
    AFTER INSERT ON auth.users
    FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- ------------------------------------------------------------------------------
-- 7. ENABLE REALTIME REPLICATION FOR WEBSOCKET LIVE DASHBOARD UPDATES
-- ------------------------------------------------------------------------------
ALTER PUBLICATION supabase_realtime ADD TABLE public.production_data;
ALTER PUBLICATION supabase_realtime ADD TABLE public.quality_data;
ALTER PUBLICATION supabase_realtime ADD TABLE public.downtime_data;
ALTER PUBLICATION supabase_realtime ADD TABLE public.profiles;

-- ------------------------------------------------------------------------------
-- 8. ROW LEVEL SECURITY (RLS) POLICIES
-- ------------------------------------------------------------------------------
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.production_data ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.quality_data ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.downtime_data ENABLE ROW LEVEL SECURITY;

-- Helper function to check user role from profiles
CREATE OR REPLACE FUNCTION public.get_user_role(uid UUID)
RETURNS TEXT AS $$
    SELECT role FROM public.profiles WHERE user_id = uid AND active = TRUE LIMIT 1;
$$ LANGUAGE sql SECURITY DEFINER STABLE;

-- PROFILES RLS
CREATE POLICY "Profiles are viewable by authenticated users"
    ON public.profiles FOR SELECT
    TO authenticated
    USING (true);

CREATE POLICY "Admins can manage profiles"
    ON public.profiles FOR ALL
    TO authenticated
    USING (public.get_user_role(auth.uid()) = 'Admin');

-- PRODUCTION_DATA RLS
CREATE POLICY "Production data viewable by all authenticated users"
    ON public.production_data FOR SELECT
    TO authenticated
    USING (is_deleted = FALSE OR public.get_user_role(auth.uid()) = 'Admin');

CREATE POLICY "Production and Admin can insert production data"
    ON public.production_data FOR INSERT
    TO authenticated
    WITH CHECK (public.get_user_role(auth.uid()) IN ('Production', 'Admin'));

CREATE POLICY "Production and Admin can update production data"
    ON public.production_data FOR UPDATE
    TO authenticated
    USING (public.get_user_role(auth.uid()) IN ('Production', 'Admin'));

CREATE POLICY "Admins can delete production data"
    ON public.production_data FOR DELETE
    TO authenticated
    USING (public.get_user_role(auth.uid()) = 'Admin');

-- QUALITY_DATA RLS
CREATE POLICY "Quality data viewable by all authenticated users"
    ON public.quality_data FOR SELECT
    TO authenticated
    USING (is_deleted = FALSE OR public.get_user_role(auth.uid()) = 'Admin');

CREATE POLICY "Quality and Admin can insert quality data"
    ON public.quality_data FOR INSERT
    TO authenticated
    WITH CHECK (public.get_user_role(auth.uid()) IN ('Quality', 'Admin', 'Production'));

CREATE POLICY "Quality and Admin can update quality data"
    ON public.quality_data FOR UPDATE
    TO authenticated
    USING (public.get_user_role(auth.uid()) IN ('Quality', 'Admin'));

CREATE POLICY "Admins can delete quality data"
    ON public.quality_data FOR DELETE
    TO authenticated
    USING (public.get_user_role(auth.uid()) = 'Admin');

-- DOWNTIME_DATA RLS
CREATE POLICY "Downtime data viewable by all authenticated users"
    ON public.downtime_data FOR SELECT
    TO authenticated
    USING (is_deleted = FALSE OR public.get_user_role(auth.uid()) = 'Admin');

CREATE POLICY "Maintenance, Production and Admin can insert downtime data"
    ON public.downtime_data FOR INSERT
    TO authenticated
    WITH CHECK (public.get_user_role(auth.uid()) IN ('Maintenance', 'Production', 'Admin'));

CREATE POLICY "Maintenance and Admin can update downtime data"
    ON public.downtime_data FOR UPDATE
    TO authenticated
    USING (public.get_user_role(auth.uid()) IN ('Maintenance', 'Admin'));

CREATE POLICY "Admins can delete downtime data"
    ON public.downtime_data FOR DELETE
    TO authenticated
    USING (public.get_user_role(auth.uid()) = 'Admin');
