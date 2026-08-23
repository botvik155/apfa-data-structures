# make        -> build both programs
# make web    -> build and start the web version
# make cli    -> build and start the console version

all: hostel server

hostel: src/console.cpp src/hostel.h
	g++ src/console.cpp -o hostel

server: src/server.cpp src/hostel.h
	g++ src/server.cpp -o server

web: server
	./server

cli: hostel
	./hostel

clean:
	rm -f hostel server
