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

# regenerate the PDF report from docs/report.html (needs google-chrome)
report:
	google-chrome --headless --disable-gpu --no-pdf-header-footer \
	  --print-to-pdf=docs/DS_Project_2_Report.pdf docs/report.html
