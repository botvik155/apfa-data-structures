// ===================================================
//  Web version -> open http://localhost:8080
//  A very small web server that shows the same
//  array + stack + string code inside an HTML page.
//
//  Build : g++ src/server.cpp -o server
//  Run   : ./server
// ===================================================

#include <iostream>
#include <fstream>
#include <sstream>
#include <cstring>
#include <unistd.h>
#include <netinet/in.h>
#include "hostel.h"

const int PORT = 8080;

// ---------- read a whole file into a string ----------
string readFile(string path) {
    ifstream f(path);
    stringstream ss;
    ss << f.rdbuf();
    return ss.str();
}

// ---------- turn "Ram%20Kumar" / "Ram+Kumar" into "Ram Kumar" ----------
string decode(string s) {
    string out = "";
    for (int i = 0; i < (int)s.size(); i++) {
        if (s[i] == '+') {
            out += ' ';
        } else if (s[i] == '%' && i + 2 < (int)s.size()) {
            string hex = s.substr(i + 1, 2);
            out += (char)strtol(hex.c_str(), NULL, 16);
            i += 2;
        } else {
            out += s[i];
        }
    }
    return out;
}

// ---------- get one value out of "roll=1&name=Ram&room=5" ----------
string getValue(string data, string key) {
    key = key + "=";
    size_t pos = data.find(key);
    if (pos == string::npos) return "";
    size_t start = pos + key.size();
    size_t end   = data.find("&", start);
    if (end == string::npos) end = data.size();
    return decode(data.substr(start, end - start));
}

// ---------- build the HTML page ----------
string makePage(string search) {
    string html = readFile("web/index.html");

    // ---- table rows from the ARRAY ----
    string rows = "";
    int shown = 0;
    for (int i = 0; i < count_; i++) {
        if (!matches(i, search)) continue;      // STRING search
        shown++;
        rows += "<tr><td>" + to_string(students[i].roll) + "</td>";
        rows += "<td>" + students[i].name + "</td>";
        rows += "<td>" + students[i].course + "</td>";
        rows += "<td>" + to_string(students[i].room) + "</td>";
        rows += "<td><a class='del' href='/delete?roll="
                + to_string(students[i].roll) + "'>delete</a></td></tr>";
    }
    if (shown == 0)
        rows = "<tr><td colspan='5' class='empty'>No students to show</td></tr>";

    // ---- the STACK of deleted students ----
    string stackList = "";
    if (stackEmpty()) {
        stackList = "<li class='empty'>stack is empty</li>";
    } else {
        for (int i = top_; i >= 0; i--) {           // top first
            stackList += "<li>";
            if (i == top_) stackList += "<b>TOP &rarr; </b>";
            stackList += stack_[i].name + " (" + to_string(stack_[i].roll) + ")";
            stackList += "</li>";
        }
    }

    // ---- put our data inside the HTML template ----
    string marks[4]  = { "{{ROWS}}", "{{STACK}}", "{{COUNT}}", "{{SEARCH}}" };
    string values[4] = { rows, stackList, to_string(count_), search };
    for (int i = 0; i < 4; i++) {
        size_t p = html.find(marks[i]);
        if (p != string::npos) html.replace(p, marks[i].size(), values[i]);
    }
    return html;
}

int main() {
    // ---- some sample data so the page is not empty ----
    addStudent(101, "Ravi Sharma",  "BCA", 12);
    addStudent(102, "Anita Verma",  "BSc", 8);
    addStudent(103, "Karan Mehta",  "BTech", 15);

    // ---- normal socket setup ----
    int server = socket(AF_INET, SOCK_STREAM, 0);
    int yes = 1;
    setsockopt(server, SOL_SOCKET, SO_REUSEADDR, &yes, sizeof(yes));

    sockaddr_in addr;
    addr.sin_family      = AF_INET;
    addr.sin_addr.s_addr = INADDR_ANY;
    addr.sin_port        = htons(PORT);

    if (bind(server, (sockaddr*)&addr, sizeof(addr)) < 0) {
        cout << "Port " << PORT << " is busy.\n";
        return 1;
    }
    listen(server, 5);
    cout << "Server running -> http://localhost:" << PORT << "\n";

    while (true) {
        int client = accept(server, NULL, NULL);
        if (client < 0) continue;

        char buffer[4096] = {0};
        read(client, buffer, 4095);
        string request = buffer;

        // first line looks like:  GET /delete?roll=101 HTTP/1.1
        string path = "/";
        size_t s1 = request.find(" ");
        size_t s2 = request.find(" ", s1 + 1);
        if (s1 != string::npos && s2 != string::npos)
            path = request.substr(s1 + 1, s2 - s1 - 1);

        // the form data is after the blank line (for POST)
        string body = "";
        size_t b = request.find("\r\n\r\n");
        if (b != string::npos) body = request.substr(b + 4);

        string response;

        // ---------- the CSS file ----------
        if (path == "/style.css") {
            string css = readFile("web/style.css");
            response = "HTTP/1.1 200 OK\r\nContent-Type: text/css\r\n"
                       "Content-Length: " + to_string(css.size()) +
                       "\r\n\r\n" + css;
        }
        // ---------- ADD  (array insert) ----------
        else if (path == "/add") {
            int    roll   = atoi(getValue(body, "roll").c_str());
            string name   = getValue(body, "name");
            string course = getValue(body, "course");
            int    room   = atoi(getValue(body, "room").c_str());

            if (roll > 0 && name != "")
                addStudent(roll, name, course, room);

            response = "HTTP/1.1 303 See Other\r\nLocation: /\r\n\r\n";
        }
        // ---------- DELETE  (array delete + stack push) ----------
        else if (path.rfind("/delete", 0) == 0) {
            int roll = atoi(getValue(path, "roll").c_str());
            deleteStudent(roll);
            response = "HTTP/1.1 303 See Other\r\nLocation: /\r\n\r\n";
        }
        // ---------- UNDO  (stack pop) ----------
        else if (path == "/undo") {
            undoDelete();
            response = "HTTP/1.1 303 See Other\r\nLocation: /\r\n\r\n";
        }
        // ---------- SORT  (string compare) ----------
        else if (path == "/sort") {
            sortByName();
            response = "HTTP/1.1 303 See Other\r\nLocation: /\r\n\r\n";
        }
        // ---------- the main page (also handles search) ----------
        else {
            string search = "";
            if (path.find("q=") != string::npos)
                search = getValue(path, "q");

            string page = makePage(search);
            response = "HTTP/1.1 200 OK\r\nContent-Type: text/html\r\n"
                       "Content-Length: " + to_string(page.size()) +
                       "\r\n\r\n" + page;
        }

        write(client, response.c_str(), response.size());
        close(client);
    }
    close(server);
    return 0;
}
