CXX = g++

CXXFLAGS = -std=gnu++98 -O0 -fno-exceptions -fno-rtti -ffunction-sections -fdata-sections 
LIBS = -L. -lbass -lcomctl32 -lgdi32
TARGET = Album95.exe
OBJS = main.o

all: $(TARGET)

$(TARGET): $(OBJS)
	$(CXX) $(OBJS) -o $(TARGET) -mwindows \
	  -static-libgcc -static-libstdc++ \
	  -Wl,--gc-sections -s \
	  $(LIBS)

%.o: %.cpp
	$(CXX) $(CXXFLAGS) -c $<

clean:
	del /q *.o 2>nul
	del /q $(TARGET) 2>nul