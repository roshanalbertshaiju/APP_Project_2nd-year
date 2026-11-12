-- ============================================================================
-- ExamOpt: Intelligent Examination Scheduling & Optimization System
-- SRM Institute of Science and Technology, Tiruchirappalli Campus
-- Database Schema Definition (DDL)
-- ============================================================================

PRAGMA foreign_keys = ON;

-- ----------------------------------------------------------------------------
-- Table: students
-- Stores registered student details across academic departments and semesters
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS students (
    student_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    department TEXT NOT NULL,
    semester INTEGER DEFAULT 1
);

-- ----------------------------------------------------------------------------
-- Table: subjects
-- Academic courses requiring examination scheduling
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS subjects (
    subject_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    department TEXT NOT NULL,
    semester INTEGER DEFAULT 1,
    duration_minutes INTEGER DEFAULT 180
);

-- ----------------------------------------------------------------------------
-- Table: enrollments
-- Associative entity mapping student course registrations (M:N relationship)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS enrollments (
    student_id TEXT NOT NULL,
    subject_id TEXT NOT NULL,
    PRIMARY KEY (student_id, subject_id),
    FOREIGN KEY (student_id) REFERENCES students(student_id) ON DELETE CASCADE,
    FOREIGN KEY (subject_id) REFERENCES subjects(subject_id) ON DELETE CASCADE
);

-- ----------------------------------------------------------------------------
-- Table: faculty
-- Faculty members designated as invigilators / examination proctors
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS faculty (
    faculty_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    department TEXT NOT NULL,
    is_available INTEGER DEFAULT 1
);

-- ----------------------------------------------------------------------------
-- Table: rooms
-- Physical venues/examination halls with designated seating capacities
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS rooms (
    room_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    capacity INTEGER NOT NULL,
    is_available INTEGER DEFAULT 1
);

-- ----------------------------------------------------------------------------
-- Table: timeslots
-- Designated examination scheduling windows (Date, Start Time, End Time)
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS timeslots (
    slot_id TEXT PRIMARY KEY,
    date TEXT NOT NULL,
    start_time TEXT NOT NULL,
    end_time TEXT NOT NULL,
    is_available INTEGER DEFAULT 1
);

-- ----------------------------------------------------------------------------
-- Table: exams
-- Core scheduled examination entities with foreign key links
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS exams (
    exam_id TEXT PRIMARY KEY,
    subject_id TEXT NOT NULL,
    slot_id TEXT,
    room_id TEXT,
    faculty_id TEXT,
    student_count INTEGER DEFAULT 0,
    status TEXT DEFAULT 'Unscheduled',
    status_reason TEXT DEFAULT '',
    FOREIGN KEY (subject_id) REFERENCES subjects(subject_id) ON DELETE CASCADE,
    FOREIGN KEY (slot_id) REFERENCES timeslots(slot_id) ON DELETE SET NULL,
    FOREIGN KEY (room_id) REFERENCES rooms(room_id) ON DELETE SET NULL,
    FOREIGN KEY (faculty_id) REFERENCES faculty(faculty_id) ON DELETE SET NULL
);

-- ----------------------------------------------------------------------------
-- Performance Indexes for Query Optimization & High-Speed Joins
-- ----------------------------------------------------------------------------
CREATE INDEX IF NOT EXISTS idx_enrollments_subject ON enrollments(subject_id);
CREATE INDEX IF NOT EXISTS idx_enrollments_student ON enrollments(student_id);
CREATE INDEX IF NOT EXISTS idx_exams_slot ON exams(slot_id);
CREATE INDEX IF NOT EXISTS idx_exams_room ON exams(room_id);
CREATE INDEX IF NOT EXISTS idx_exams_faculty ON exams(faculty_id);
