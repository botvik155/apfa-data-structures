// ===================================================
//  Hostel Student Management System
//  Data structures used:  ARRAY , STACK , STRING
// ===================================================

#ifndef HOSTEL_H
#define HOSTEL_H

#include <string>
using namespace std;

const int MAX = 100;

// ---------- one student record ----------
struct Student {
    int    roll;
    string name;      // STRING
    string course;    // STRING
    int    room;
};

// ---------- ARRAY : stores all students ----------
Student students[MAX];
int count_ = 0;

// ---------- STACK : stores deleted students (for undo) ----------
Student stack_[MAX];
int top_ = -1;


// ================= ARRAY functions =================

// search a roll number, return its position (-1 if not found)
int findStudent(int roll) {
    for (int i = 0; i < count_; i++)
        if (students[i].roll == roll)
            return i;
    return -1;
}

// add a student at the end of the array
bool addStudent(int roll, string name, string course, int room) {
    if (count_ >= MAX)          return false;   // array is full
    if (findStudent(roll) != -1) return false;   // roll already exists

    students[count_].roll   = roll;
    students[count_].name   = name;
    students[count_].course = course;
    students[count_].room   = room;
    count_++;
    return true;
}


// ================= STACK functions =================

void push(Student s) {              // put on top of stack
    if (top_ < MAX - 1) {
        top_++;
        stack_[top_] = s;
    }
}

Student pop() {                     // take from top of stack
    Student s = stack_[top_];
    top_--;
    return s;
}

bool stackEmpty() {
    return top_ == -1;
}


// ============ ARRAY delete + STACK push ============

bool deleteStudent(int roll) {
    int i = findStudent(roll);
    if (i == -1) return false;

    push(students[i]);                    // remember it, so we can undo

    for (int j = i; j < count_ - 1; j++)   // shift the array left
        students[j] = students[j + 1];
    count_--;
    return true;
}

// undo = pop from stack and put back in array
bool undoDelete() {
    if (stackEmpty()) return false;
    Student s = pop();
    return addStudent(s.roll, s.name, s.course, s.room);
}


// ================ STRING functions ================

// make a string small letters
string toLower(string s) {
    for (int i = 0; i < (int)s.size(); i++)
        if (s[i] >= 'A' && s[i] <= 'Z')
            s[i] = s[i] + 32;
    return s;
}

// does 'text' contain 'key' ?  (ignoring capital letters)
bool contains(string text, string key) {
    if (key == "") return true;
    return toLower(text).find(toLower(key)) != string::npos;
}

// does student number i match the search word ?
bool matches(int i, string key) {
    return contains(students[i].name, key) ||
           contains(students[i].course, key) ||
           contains(to_string(students[i].roll), key);
}

// sort the array by name (bubble sort, compares strings)
void sortByName() {
    for (int i = 0; i < count_ - 1; i++)
        for (int j = 0; j < count_ - 1 - i; j++)
            if (toLower(students[j].name) > toLower(students[j + 1].name)) {
                Student t       = students[j];
                students[j]     = students[j + 1];
                students[j + 1] = t;
            }
}

#endif
