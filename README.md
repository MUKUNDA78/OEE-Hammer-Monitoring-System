# 🏭 GRS Engineering Pvt Ltd Unit 1 - OEE & Quality Monitoring Cloud System

An industrial-grade, multi-user 5-hammer fleet efficiency, OEE monitoring, and Quality Activity Analytics web application powered by **Supabase PostgreSQL**, WebSockets Realtime, Supabase Authentication, and Row Level Security (RLS).

![OEE System Banner](https://img.shields.io/badge/System-Supabase_Realtime_Cloud_OEE-16a34a?style=for-the-badge&logo=supabase)
![Security](https://img.shields.io/badge/Security-RLS_PostgreSQL_Auth-0284c7?style=for-the-badge&logo=postgresql)
![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)

---

## 🌟 Architecture & Highlights

1. **Central Cloud Database (Supabase PostgreSQL)**:
   - Single source of truth across desktop, laptop, tablet, and mobile devices.
   - Tables: `production_data`, `quality_data`, `downtime_data`, and `profiles`.
   - Soft-delete support (`is_deleted`, `deleted_by`, `deleted_at`) and audit timestamps (`created_by`, `created_at`, `updated_at`).
   - Maintains **Rework** and **Rejection** quantities separately.

2. **Real-time WebSockets & Presence**:
   - Near real-time bi-directional WebSocket data sync across all connected users without browser refresh.
   - Real-time Connection Pill: 🟢 `Live Database Connected` | 🔴 `Offline / Reconnecting`.
   - Real-time Presence Indicator: Displays online user count (e.g. `🟢 4 Users Online`) with role/department breakdown popover.

3. **Supabase Auth & Role-Based Access Control (RLS)**:
   - 6 System Roles:
     - `Admin`: Full access, user management, complete database control.
     - `Quality`: Add/edit quality records, view quality activity monitor, production, and downtime.
     - `Production`: Add/edit production & in-process shift records.
     - `Maintenance`: Add/edit downtime records.
     - `Management`: Read-only access to all dashboards and reports.
     - `Viewer`: Read-only access.
   - Strict Row Level Security (RLS) policies prevent unauthorized modifications.

4. **Quality Activity Monitor**:
   - Dedicated tab with strict filter isolation (`Month`, `Inspection Stage`: `In-Process`, `Final Inspection`, `MPI`, `Part Number`, `Defect Reason`).
   - Metrics: Total Inspected, Total Rework, Total Rejection, Rework %, Rejection %.
   - Part-Wise Rework Analysis & Part-Wise Rejection Analysis tables.
   - Reason-Wise Defect Pareto Chart.
   - Monthly Rework & Rejection Trend Charts.
   - Top 10 Rework Reasons & Top 10 Rejection Reasons.

5. **Excel Import Validation & Categorization**:
   - Validates dates, part numbers, quantities, inspection stages, and downtime categories.
   - Compares rows against existing database records to detect duplicates.
   - Categorizes rows into **`New Valid Records`**, **`Duplicate Records`**, and **`Error / Invalid Records`** prior to final insertion.

---

## ⚡ Supabase Setup & Execution Instructions

### Step 1: Create a Supabase Project
1. Log in to [Supabase Dashboard](https://database.new).
2. Click **New Project** -> Enter Name: `grs-unit1-oee-quality-monitor`.
3. Set Database Password & select Region -> Click **Create New Project**.
4. Navigate to **Project Settings** -> **API**:
   - Copy **Project URL** (`https://xyz.supabase.co`)
   - Copy **Anon / Public API Key** (`eyJhbGci...`)

### Step 2: Execute Database DDL Schema Script
1. In Supabase Dashboard, open the **SQL Editor** (`/project/_/sql`).
2. Open the included [`supabase_schema.sql`](file:///C:/Users/mukun/.gemini/antigravity/scratch/oee-hammer-monitor/supabase_schema.sql) file.
3. Paste the contents into the SQL Editor and click **Run**.
4. This creates:
   - `profiles`, `production_data`, `quality_data`, and `downtime_data` tables.
   - Indexes, triggers, updated_at handlers, and automatic profile creation triggers.
   - Enables Realtime publications (`ALTER PUBLICATION supabase_realtime ADD TABLE ...`).
   - Applies Row Level Security (RLS) policies for all 6 roles.

### Step 3: Configure Environment Variables or UI Connection Modal
- For local / static hosting, click **`DB Config`** in the application header and paste your `Supabase Project URL` and `Supabase Anon Key`.
- For Vite / bundler builds, set environment variables:
  ```env
  VITE_SUPABASE_URL=https://xyz.supabase.co
  VITE_SUPABASE_ANON_KEY=eyJhbGciOiJIUzI1NiIsInR...
  ```

---

## 👤 User Creation Guide (Creating Users for All 6 Roles)

### Creating the First Admin User:
1. Open the application -> Click **`Login / Register`** -> Select **`Create Account`**.
2. Register:
   - Email: `admin@grsengineering.com`
   - Password: `YourSecurePassword123`
   - Employee Name: `Plant Admin`
   - Department: `Administration`
   - Role: `Admin`
3. Click **Create Supabase Account**.

### Adding Operational Users:
Repeat the signup process or create users directly in Supabase **Authentication** -> **Users**:
- **Quality User**: Role = `Quality`, Dept = `Quality` (e.g. `quality@grsengineering.com`)
- **Production User**: Role = `Production`, Dept = `Production` (e.g. `prod@grsengineering.com`)
- **Maintenance User**: Role = `Maintenance`, Dept = `Maintenance` (e.g. `maint@grsengineering.com`)
- **Management User**: Role = `Management`, Dept = `Management` (e.g. `management@grsengineering.com`)
- **Viewer User**: Role = `Viewer`, Dept = `Operations` (e.g. `viewer@grsengineering.com`)

---

## 🧪 Simultaneous Multi-User Testing Procedure

To verify real-time bi-directional sync across multiple devices/computers:

1. **Window / Computer 1 (Operator)**:
   - Log in as `Production` or `Quality`.
   - Open Manual Entry Form or Excel Import and insert a shift entry for `1.5 Ton Hammer` (Part `W1#164`).
   - Click **Save**.

2. **Window / Computer 2 (Management / Mobile Phone)**:
   - Open the application link on another computer or phone as `Management`.
   - Notice that the **Overview Dashboard**, **Quality Activity Monitor**, and **Shift Table** automatically receive the new record via WebSockets **without manually refreshing the page**!
   - Observe the presence indicator displaying `🟢 2 Users Online`.

---

## 🌐 Deployment Instructions

### Deployment to Netlify / Vercel / Custom Company Domain:
1. Push repository to GitHub.
2. Connect repository to Netlify or Vercel.
3. Set Build Environment Variables:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_ANON_KEY`
4. Access via company domain (e.g. `https://quality-monitor.grsengineering.com`).

---

## 📜 Repository Files

```
oee-hammer-monitor/
├── index.html           # HTML5 Layout with Quality Monitor & Auth Modals
├── styles.css           # Industrial Mint Theme, Presence & Quality Monitor CSS
├── app.js               # Supabase Engine, Auth, Realtime, Quality Analytics
├── supabase_schema.sql  # Supabase PostgreSQL DDL, Triggers, Realtime & RLS Policies
├── grs_logo.png         # Official GRS Engineering Pvt Ltd Logo
└── README.md            # Complete Setup & User Guide
```
