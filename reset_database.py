#!/usr/bin/env python3
"""
Database Reset Script
Run this script to fix database lock issues
"""

import sqlite3
import os
import sys
import time

def reset_database():
    """Reset the database to fix lock issues"""

    print("🔧 Database Reset Script")
    print("=" * 50)

    # Check if database exists
    if not os.path.exists('school.db'):
        print("❌ Database file 'school.db' not found!")
        print("Please make sure you're in the correct directory.")
        return False

    print(f"✅ Found database file: {os.path.getsize('school.db')} bytes")

    try:
        # Try to backup the current database
        backup_file = f'school_backup_{int(time.time())}.db'
        import shutil
        shutil.copy('school.db', backup_file)
        print(f"✅ Database backup created: {backup_file}")

        # Try to connect with timeout
        print("🔄 Attempting to connect to database...")
        conn = sqlite3.connect('school.db', timeout=10)

        # Test basic operations
        cursor = conn.cursor()

        # Check tables
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchone()
        if tables:
            print(f"✅ Database contains tables: {[t[0] for t in cursor.fetchall()]}")
        else:
            print("⚠️  No tables found in database")

        # Check student count
        try:
            cursor.execute("SELECT COUNT(*) FROM students")
            student_count = cursor.fetchone()[0]
            print(f"✅ Students in database: {student_count}")
        except:
            print("⚠️  Students table may not exist or be accessible")

        # Check attendance count
        try:
            cursor.execute("SELECT COUNT(*) FROM attendance")
            attendance_count = cursor.fetchone()[0]
            print(f"✅ Attendance records in database: {attendance_count}")
        except:
            print("⚠️  Attendance table may not exist or be accessible")

        conn.close()
        print("✅ Database connection successful!")
        print("✅ Database is working correctly!")

        # Check if IT attendance data exists
        print("\n🔍 Checking IT attendance data...")
        conn = sqlite3.connect('school.db')
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM attendance WHERE rollno LIKE '%IT%' OR rollno LIKE '3%' OR rollno LIKE '4%'")
        it_attendance = cursor.fetchone()[0]
        print(f"📊 IT attendance records: {it_attendance}")

        cursor.execute("SELECT DISTINCT rollno FROM attendance WHERE rollno LIKE '%IT%' OR rollno LIKE '3%' OR rollno LIKE '4%' LIMIT 5")
        it_rollnos = cursor.fetchall()
        if it_rollnos:
            print("✅ Sample IT roll numbers with attendance:")
            for rollno in it_rollnos:
                print(f"   {rollno[0]}")
        else:
            print("❌ No IT attendance data found")

        conn.close()

        return True

    except sqlite3.OperationalError as e:
        if 'database is locked' in str(e):
            print("❌ Database is locked by another process")
            print("\n🔧 Solutions:")
            print("1. Close all Python processes and IDEs")
            print("2. Wait 30 seconds and try again")
            print("3. Restart your computer")
            print("4. Delete the database and recreate it")

            # Try to force unlock by creating a new connection
            try:
                print("🔄 Attempting to force unlock...")
                conn = sqlite3.connect('school.db', timeout=1)
                conn.close()
                print("✅ Force unlock attempt completed")
            except:
                pass

            return False
        else:
            print(f"❌ Database error: {e}")
            return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False

def create_fresh_database():
    """Create a fresh database if needed"""

    print("\n🔄 Creating fresh database...")

    # Remove the old database
    if os.path.exists('school.db'):
        os.remove('school.db')
        print("✅ Removed old database")

    # Create new database with basic structure
    conn = sqlite3.connect('school.db')
    cursor = conn.cursor()

    # Create tables
    cursor.execute('''
        CREATE TABLE students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            reg_no TEXT,
            rollno TEXT UNIQUE,
            name TEXT,
            dob TEXT,
            gender TEXT,
            aadhar TEXT,
            student_mobile TEXT,
            blood_group TEXT,
            parent_name TEXT,
            parent_mobile TEXT,
            address TEXT,
            nationality TEXT,
            religion TEXT,
            community TEXT,
            caste TEXT,
            day_scholar_or_hosteller TEXT,
            current_semester TEXT,
            seat_type TEXT,
            quota_type TEXT,
            email TEXT,
            pmss TEXT,
            remarks TEXT,
            bus_no TEXT,
            hosteller_room_no TEXT,
            outside_staying_address TEXT,
            owner_ph_no TEXT,
            user_id TEXT UNIQUE,
            password_hash TEXT,
            password_plain TEXT,
            extra_json TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE attendance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rollno TEXT,
            reg_no TEXT,
            date TEXT,
            status TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE teachers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            teacher_name TEXT,
            department TEXT,
            user_id TEXT UNIQUE,
            pass_hash TEXT,
            pass_plain TEXT,
            qualification TEXT,
            experience TEXT,
            subject TEXT,
            address TEXT,
            date_of_joining TEXT,
            salary TEXT,
            extra_json TEXT,
            role TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE courses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            course_name TEXT,
            course_code TEXT UNIQUE,
            drive_link TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE out_passes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_role TEXT,
            requester_user_id TEXT,
            requester_name TEXT,
            rollno TEXT,
            department TEXT,
            pass_type TEXT,
            reason TEXT,
            from_datetime TEXT,
            to_datetime TEXT,
            od_duration TEXT,
            od_days INTEGER,
            other_hours TEXT,
            status TEXT DEFAULT 'pending',
            approver_user_id TEXT,
            remarks TEXT,
            advisor_status TEXT DEFAULT 'pending',
            hod_status TEXT DEFAULT 'pending',
            advisor_user_id TEXT,
            advisor_remarks TEXT,
            hod_user_id TEXT,
            hod_remarks TEXT,
            created_at INTEGER,
            updated_at INTEGER
        )
    ''')

    conn.commit()
    conn.close()

    print("✅ Fresh database created successfully!")
    return True

if __name__ == "__main__":
    print("Starting database reset process...")

    # Try to reset existing database first
    if reset_database():
        print("\n🎉 Database is working correctly!")
        print("You can now run your Flask application.")
    else:
        print("\n🔄 Database reset failed, creating fresh database...")
        if create_fresh_database():
            print("\n🎉 Fresh database created!")
            print("You can now run your Flask application.")
        else:
            print("\n❌ Failed to create database")
            sys.exit(1)
