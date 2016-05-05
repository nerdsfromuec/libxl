#include <iostream>
#include <stdio.h>
#include <conio.h>
#include <string.h>
#include "libxl.h"

using namespace std;
using namespace libxl;

class Bus
{
public:
int initRoutesAndBuses();
int getBus();
int getRoute();
};

int Bus::initRoutesAndBuses() 
{	
    Book* book = xlCreateBook();
    if(book)
    {
        Sheet* sheet = book->addSheet("Bus_Route");
        if(sheet)
        {
            sheet->writeStr(2, 1, "All_Routes              ");
            sheet->writeStr(2, 2, " Bus_Id's(Bus_Number's)-------->");

            // Bus Number's.

            sheet->writeStr(3, 3, "1");
            sheet->writeStr(3, 4, "1A");
            sheet->writeStr(3, 5, "2");
            sheet->writeStr(3, 6, "4");
            sheet->writeStr(3, 7, "4A");
            sheet->writeStr(3, 8, "6");
            sheet->writeStr(3, 9, "8");
            sheet->writeStr(3, 10, "9");

           // All Stations  with Corresponding Bus

            sheet->writeStr(5, 1, "Abilasha Colony");
            sheet->writeNum(5,9,1);
            
            sheet->writeStr(6, 1, "Ahmad Nagar");
            sheet->writeNum(6,7,1);

            sheet->writeStr(7, 1, "Ankpat Marg Tiraha");
            sheet->writeNum(7,3,1);
            sheet->writeNum(7,4,1);
            sheet->writeNum(7,5,1);

            sheet->writeStr(8, 1, "Atlas Chouraha");
            sheet->writeNum(8,9,1);

            sheet->writeStr(9, 1, "Bada Hospital");
            sheet->writeNum(9,6,1);
            sheet->writeNum(9,7,1);

            sheet->writeStr(10,1, "Begam Bag");
            sheet->writeNum(10,10,1);

            sheet->writeStr(11,1, "Begam Bag Road");
            sheet->writeNum(11,9,1);


            sheet->writeStr(12,1, "Bhairavgarh");
            sheet->writeNum(12,5,1);


            sheet->writeStr(13,1, "Bhairavgarh Chouraha");
            sheet->writeNum(13,3,1);
            sheet->writeNum(13,4,1);

            sheet->writeStr(14,1, "Bharat Puri");
            sheet->writeNum(14,2,0);

            sheet->writeStr(15,1, "BheravGard Jail Tiraha");
            sheet->writeNum(15,5,1);

            sheet->writeStr(16,1, "Bhima Hospital");
            sheet->writeNum(15,6,1);
            sheet->writeNum(15,7,1);

            sheet->writeStr(17,1, "Birla Hospital");
            sheet->writeNum(17,2,1);

            sheet->writeStr(18,1, "Bhudhwaria");
            sheet->writeNum(18,4,1);
            sheet->writeNum(18,5,1);


            sheet->writeStr(19,1, "Bus Stand");
            sheet->writeNum(19,3,1);
            sheet->writeNum(19,4,1);
            sheet->writeNum(19,5,1);
            sheet->writeNum(19,9,1);
            sheet->writeNum(19,10,1);

            sheet->writeStr(20,1, "Chak kamaed");
            sheet->writeNum(20,7,1);

            sheet->writeStr(21,1, "Chmunda Chouraha");
            sheet->writeNum(21,3,1);
            sheet->writeNum(21,4,1);
            sheet->writeNum(21,5,1);
            sheet->writeNum(21,6,1);
            sheet->writeNum(21,7,1);
            sheet->writeNum(21,9,1);
            sheet->writeNum(21,10,1);


            sheet->writeStr(22,1, "Chatri Chowk");
            sheet->writeNum(22,9,1);

            sheet->writeStr(23,1, "Chintaman jawasia");
            sheet->writeNum(23,10,1);

            sheet->writeStr(24,1, "City Bus Depot");
            sheet->writeNum(24,8,1);

            sheet->writeStr(25,1, "Daulat Ganj");
            sheet->writeNum(25,3,1);
            sheet->writeNum(25,4,1);
            sheet->writeNum(25,5,1);
            sheet->writeNum(25,10,1);

            sheet->writeStr(26,1, "Daulat Ganj Chouraha");
            sheet->writeNum(26,9,1);

            sheet->writeStr(27,1, "Desai Nagar Tiraha");
            sheet->writeNum(27,9,1);

            sheet->writeStr(28,1, "Dinesh Petrol Pump");
            sheet->writeNum(28,3,1);
            sheet->writeNum(28,4,1);
            sheet->writeNum(28,7,1);

            sheet->writeStr(29,1, "Dipti Vihar");
            sheet->writeNum(29,3,1);

            sheet->writeStr(30,1, "Durga Plaza");
            sheet->writeNum(30,6,1);
            sheet->writeNum(30,10,1);

            sheet->writeStr(31,1, "Gail Colony");
            sheet->writeNum(31,9,1);

            sheet->writeStr(32,1, "Ghas Mandi Churaha");
            sheet->writeNum(32,9,1);

            sheet->writeStr(33,1, "Gopal Pura");
            sheet->writeNum(33,8,1);

            sheet->writeStr(34,1, "Hamukhedi");
            sheet->writeNum(34,9,1);

            sheet->writeStr(35,1, "Harsidhi Mandir");
            sheet->writeNum(35,9,1);

            sheet->writeStr(36,1, "Indira Nagar Tiraha");
            sheet->writeNum(36,6,1);
            sheet->writeNum(36,7,1);

            sheet->writeStr(37,1, "Indore Gate");
            sheet->writeNum(37,3,1);
            sheet->writeNum(37,4,1);
            sheet->writeNum(37,5,1);
            sheet->writeNum(37,9,1);
            sheet->writeNum(37,10,1);

            sheet->writeStr(38,1, "ISKON Temple Choraha");
            sheet->writeNum(38,6,1);

            sheet->writeStr(39,1, "Jail Tiraha");
            sheet->writeNum(39,3,1);
            sheet->writeNum(39,4,1);

            sheet->writeStr(40,1, "Jaisingh Pura");
            sheet->writeNum(40,10,1);

            sheet->writeStr(41,1, "jantar Mantar");
            sheet->writeNum(41,10,1);

            sheet->writeStr(42,1, "Jawahar Nagar Churaha");
            sheet->writeNum(42,10,1);

            sheet->writeStr(43,1, "kala Patthar");
            sheet->writeNum(43,3,1);

            sheet->writeStr(44,1, "kaliyadeh Mahal");
            sheet->writeNum(44,3,1);


            sheet->writeStr(45,1, "kamal Colony Tiraha");
            sheet->writeNum(45,3,1);
            sheet->writeNum(45,4,1);
            sheet->writeNum(45,5,1);
                              
            sheet->writeStr(46,1, "Kanthal");
            sheet->writeNum(46,3,1);
            sheet->writeNum(46,4,1);
            sheet->writeNum(46,5,1);
            sheet->writeNum(46,9,1);

            sheet->writeStr(47,1, "Kendriya Vidhyalay");
            sheet->writeNum(47,3,1);
            sheet->writeNum(47,4,1);
            sheet->writeNum(47,6,1);

            sheet->writeStr(48,1, "Khajurwale Baba");
            sheet->writeNum(48,3,1);
            sheet->writeNum(48,7,1);

            sheet->writeStr(49,1, "Kishan Pura");
            sheet->writeNum(49,8,1);

            sheet->writeStr(50,1, "Kothi");
            sheet->writeNum(50,9,1);
            sheet->writeNum(50,10,1);

            sheet->writeStr(51,1, "Koyla Phatak");
            sheet->writeNum(51,6,1);
            sheet->writeNum(51,7,1);

            sheet->writeStr(52,1, "Mahakal Chowk/Maidan");
            sheet->writeNum(52,10,1);

            sheet->writeStr(53,1, "Mahakal Sindhi Colony");
            sheet->writeNum(53,3,1);
            sheet->writeNum(53,4,1);
            sheet->writeNum(53,7,1);

            sheet->writeStr(54,1, "Makodiya Naka");
            sheet->writeNum(54,6,1);
            sheet->writeNum(54,7,1);

            sheet->writeStr(55,1, "Mandi Tiraha");
            sheet->writeNum(55,6,1);
            sheet->writeNum(55,7,1);

            sheet->writeStr(56,1, "Maxi Road Tiraha");
            sheet->writeNum(56,8,1);

            sheet->writeStr(57,1, "Mungi Chouraha");
            sheet->writeNum(57,6,1);

            sheet->writeStr(58,1, "Nagziri");
            sheet->writeNum(58,9,1);

            sheet->writeStr(59,1, "Nanakheda");
            sheet->writeNum(59,3,1);
            sheet->writeNum(59,4,1);
            sheet->writeNum(59,7,1);
            sheet->writeNum(59,10,1);

            sheet->writeStr(60,1, "Panch Banglow");
            sheet->writeNum(60,10,1);

            sheet->writeStr(61,1, "Patel Colony");
            sheet->writeNum(61,3,1);
            sheet->writeNum(61,4,1);
            sheet->writeNum(61,5,1);

            sheet->writeStr(62,1, "Pawasa      ");
            sheet->writeNum(62,8,1);

            sheet->writeStr(63,1, "Pipe Factory");
            sheet->writeNum(63,6,1);
            sheet->writeNum(63,8,1);

            sheet->writeStr(64,1, "Pipli Naka");
            sheet->writeNum(64,3,1);
            sheet->writeNum(64,4,1);
            sheet->writeNum(64,5,1);

            sheet->writeStr(65,1, "Police Mess");
            sheet->writeNum(65,6,1);
            sheet->writeNum(65,10,1);

            sheet->writeStr(66,1, "Polythecnic Clg.");
            sheet->writeNum(66,6,1);
            sheet->writeNum(66,9,1);
            sheet->writeNum(66,10,1);

            sheet->writeStr(67,1, "Prashanti Dham");
            sheet->writeNum(67,3,1);

            sheet->writeStr(68,1, "Pushpa Mission Hospital");
            sheet->writeNum(68,6,1);

            sheet->writeStr(69,1, "Railway Crossing");
            sheet->writeNum(69,10,1);

            sheet->writeStr(70,1, "Railway Station");
            sheet->writeNum(70,3,1);
            sheet->writeNum(70,4,1);
            sheet->writeNum(70,5,1);
            sheet->writeNum(70,9,1);
            sheet->writeNum(70,10,1);

            sheet->writeStr(71,1, "R.D. Gardy Medical Clg.");
            sheet->writeNum(71,6,1);
            sheet->writeNum(71,7,1);

            sheet->writeStr(72,1, "Rishi Nagar");
            sheet->writeNum(72,6,1);

            sheet->writeStr(73,1, "Sabji Mandi");
            sheet->writeNum(73,8,1);

            sheet->writeStr(74,1, "Shankar Pura");
            sheet->writeNum(74,8,1);

            sheet->writeStr(75,1, "Sant Nagar");
            sheet->writeNum(75,3,1);
            sheet->writeNum(75,4,1);
            sheet->writeNum(75,7,1);

            sheet->writeStr(76,1, "SBI Nayi Sadak");
            sheet->writeNum(76,3,1);
            sheet->writeNum(76,4,1);
            sheet->writeNum(76,5,1);
            sheet->writeNum(76,9,1);

            sheet->writeStr(77,1, "Sethi Building Chouraha");
            sheet->writeNum(77,3,1);
            sheet->writeNum(77,4,1);
            sheet->writeNum(77,5,1);
            sheet->writeNum(77,7,1);
            sheet->writeNum(77,9,1);
            sheet->writeNum(77,10,1);

            sheet->writeStr(78,1, "Sethi Nagar");
            sheet->writeNum(78,9,1);

            sheet->writeStr(79,1, "Shahid Park");
            sheet->writeNum(79,9,1);

            sheet->writeStr(80,1, "Shani Mandir");
            sheet->writeNum(80,3,1);

            sheet->writeStr(81,1, "Shree Synthetics");
            sheet->writeNum(81,8,1);

            sheet->writeStr(82,1, "Sidhwat");
            sheet->writeNum(82,3,1);
            sheet->writeNum(82,4,1);
            sheet->writeNum(82,5,1);

            sheet->writeStr(83,1, "Sindhi Colony Chouraha");
            sheet->writeNum(83,3,1);
            sheet->writeNum(83,4,1);
            sheet->writeNum(83,7,1);


            sheet->writeStr(84,1, "Sodang Chouraha");
            sheet->writeNum(84,4,1);
            sheet->writeNum(84,5,1);

            sheet->writeStr(85,1, "Study Home School");
            sheet->writeNum(85,3,1);
            sheet->writeNum(85,4,1);
            sheet->writeNum(85,5,1);
            sheet->writeNum(85,7,1);
            sheet->writeNum(85,9,1);
            sheet->writeNum(85,10,1);

            sheet->writeStr(86,1, "Subhash Nagar");
            sheet->writeNum(86,3,1);
            sheet->writeNum(86,4,1);
            sheet->writeNum(86,7,1);

            sheet->writeStr(87,1, "Tapobhumi");
            sheet->writeNum(87,3,1);

            sheet->writeStr(88,1, "Teen Batti Chouraha");
            sheet->writeNum(88,3,1);
            sheet->writeNum(88,4,1);
            sheet->writeNum(88,6,1);
            sheet->writeNum(88,7,1);
            sheet->writeNum(88,9,1);
            sheet->writeNum(88,10,1);


            sheet->writeStr(89,1, "Teliwada Chouraha");
            sheet->writeNum(89,3,1);
            sheet->writeNum(89,5,1);
            sheet->writeNum(89,9,1);

            sheet->writeStr(90,1, "Tower");
            sheet->writeNum(90,3,1);
            sheet->writeNum(90,4,1);
            sheet->writeNum(90,5,1);
            sheet->writeNum(90,6,1);
            sheet->writeNum(90,7,1);
            sheet->writeNum(90,8,1);
            sheet->writeNum(90,9,1);
            sheet->writeNum(90,10,1);

            sheet->writeStr(91,1, "U... marg");
            sheet->writeNum(91,9,1);

            sheet->writeStr(92,1, "Union Bank Atlas Chouraha");
            sheet->writeNum(92,3,1);
            sheet->writeNum(92,4,1);
            sheet->writeNum(92,5,1);

            sheet->writeStr(93,1, "Upkeshwar Mahadev");
            sheet->writeNum(93,10,1);

            sheet->writeStr(94,1, "Veer Sawarkar Chouraha");
            sheet->writeNum(94,3,1);

            sheet->writeStr(94,1, "Vikram Vatika");
            sheet->writeNum(94,4,1);
            sheet->writeNum(94,5,1);
            sheet->writeNum(94,9,1);
            sheet->writeNum(94,10,1);


        }                   

        if(book->save("busRouteDataBase.xls")) std::cout << "\nFile busRouteDataBase.xls has been created." << std::endl;
        book->release();         
    }


    return 0;
}

int Bus::getBus()
{
	
Book* book = xlCreateBook();
int row;
char *source,*destination;
const char *temp;
static int col;
if(book)
{
if(book->load("busRouteDataBase.xls"))
{
Sheet* sheet = book->getSheet(0);

if(sheet)
{
	std::cout<<"Enter Source"<<endl;
	std::cin>>source;
	std::cout<<"Enter Destination"<<endl;
	std::cin>>destination;
	
for(row=5;row<94;row++)
{
//std::cout<<"Source="<<sheet->readStr(row,1)<<endl;
temp=sheet->readStr(row,1);
if(strcmpi(source,temp)==0){
std::cout<<"Found. Source at index="<<row+1<<endl;
break;
}
}
}
else
{
std::cout<<"Error  While Opening File."<<endl;
}
}
book->release();
}
std::cout << "\nPress any key to exit...";
_getch();
return 0;
}


int Bus::getRoute(){
return 0;
}

int main(void){

Bus *b;
int ch;

b=new Bus();

std::cout<<"Enter Choice."<<endl;
std::cout<<"1. Want To Initialize Route Table.(Y/N)"<<endl;
std::cout<<"2. Find Bus"<<endl;
std::cout<<"3. Find Route"<<endl;
std::cout<<"4. Exit"<<endl;

std::cin>>ch;

switch(ch){

case 1:
b->initRoutesAndBuses();
break;

case 2:
b->getBus();
break;

case 3:
b->getRoute();
break;

case 4:
return -1;

default:
std::cout<<"Invalid Choice:";
}
return 0;
}

