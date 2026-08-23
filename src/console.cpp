// ===================================================
//  Menu based console version (no web needed)
//  Build : g++ src/console.cpp -o hostel
//  Run   : ./hostel
// ===================================================

#include <iostream>
#include "hostel.h"

int main() {
    int choice;

    while (true) {
        cout << "\n===== HOSTEL STUDENT MANAGEMENT =====\n";
        cout << "1. Add student\n";
        cout << "2. Show all students\n";
        cout << "3. Search student\n";
        cout << "4. Delete student\n";
        cout << "5. Undo delete (stack)\n";
        cout << "6. Sort by name\n";
        cout << "7. Exit\n";
        cout << "Enter choice: ";
        cin >> choice;

        if (choice == 1) {
            int roll, room;
            string name, course;
            cout << "Roll   : "; cin >> roll;
            cout << "Name   : "; cin >> name;
            cout << "Course : "; cin >> course;
            cout << "Room   : "; cin >> room;

            if (addStudent(roll, name, course, room))
                cout << "Student added.\n";
            else
                cout << "Not added (roll already exists or array full).\n";
        }

        else if (choice == 2) {
            cout << "\nROLL\tNAME\t\tCOURSE\t\tROOM\n";
            for (int i = 0; i < count_; i++)
                cout << students[i].roll << "\t" << students[i].name
                     << "\t\t" << students[i].course
                     << "\t\t" << students[i].room << "\n";
            if (count_ == 0) cout << "(no students)\n";
        }

        else if (choice == 3) {
            string key;
            cout << "Search word: "; cin >> key;
            for (int i = 0; i < count_; i++)
                if (matches(i, key))
                    cout << students[i].roll << "  " << students[i].name
                         << "  " << students[i].course
                         << "  room " << students[i].room << "\n";
        }

        else if (choice == 4) {
            int roll;
            cout << "Roll to delete: "; cin >> roll;
            if (deleteStudent(roll))
                cout << "Deleted (saved in stack).\n";
            else
                cout << "Roll not found.\n";
        }

        else if (choice == 5) {
            if (undoDelete())
                cout << "Last deleted student is back.\n";
            else
                cout << "Stack is empty, nothing to undo.\n";
        }

        else if (choice == 6) {
            sortByName();
            cout << "Sorted by name.\n";
        }

        else if (choice == 7) {
            cout << "Bye\n";
            break;
        }

        else {
            cout << "Wrong choice.\n";
        }
    }
    return 0;
}
