import pandas as pd

def main():
    locations = [
        {"room_name": "Library", "block": "Main Building", "floor": "1st Floor", "room_number": "", "category": "Academic", "directions": "Take the stairs near the main entrance."},
        {"room_name": "Accounts Office", "block": "Admin Block", "floor": "Ground Floor", "room_number": "A-05", "category": "Administration", "directions": "Located right next to the reception desk."},
        {"room_name": "Chemistry Lab", "block": "Science Wing", "floor": "2nd Floor", "room_number": "S-201", "category": "Laboratory", "directions": "Second floor, right wing, opposite the staff room."},
        {"room_name": "Canteen", "block": "Campus Center", "floor": "Ground Floor", "room_number": "", "category": "Facility", "directions": "Located behind the Main Building, near the basketball court."},
        {"room_name": "Staff Room 1", "block": "A-Block", "floor": "1st Floor", "room_number": "A-102", "category": "Staff", "directions": "First floor corridor, next to the CSE department classrooms."},
        {"room_name": "HOD CSE Office", "block": "CSE Block", "floor": "Ground Floor", "room_number": "C-11", "category": "Staff", "directions": "Near the main entrance of CSE Block."},
        {"room_name": "Room 101", "block": "A-Block", "floor": "1st Floor", "room_number": "101", "category": "Classroom", "directions": "First left from the stairwell on the 1st floor."},
        {"room_name": "Room 204", "block": "A-Block", "floor": "2nd Floor", "room_number": "204", "category": "Classroom", "directions": "Down the main hall, right side."},
        {"room_name": "Computer Lab 1", "block": "CSE Block", "floor": "2nd Floor", "room_number": "CL-1", "category": "Laboratory", "directions": "Top floor, look for the big glass double doors."},
        {"room_name": "Physics Lab", "block": "Science Wing", "floor": "1st Floor", "room_number": "S-101", "category": "Laboratory", "directions": "End of the Science Wing hall on the 1st floor."}
    ]
    df = pd.DataFrame(locations)
    df.to_excel("data/locations.xlsx", index=False)
    print("Created successfully.")

if __name__ == "__main__":
    main()
