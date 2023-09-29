// JavaScript
let currentLocation;
let locations = [];

function getCurrentLocation() {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(function(position) {
            currentLocation = new google.maps.LatLng(position.coords.latitude, position.coords.longitude);
            document.getElementById("current-location").textContent = `Current Location taken`;
        });
    } else {
        alert("Geolocation is not supported by this browser.");
    }
}

// Autocomplete for base location input
const baseLocationInput = document.getElementById("base-location-input");
const baseAutocomplete = new google.maps.places.Autocomplete(baseLocationInput);

baseAutocomplete.addListener("place_changed", function() {
    const place = baseAutocomplete.getPlace();

    document.getElementById("selected-base-location-display").textContent = `Your Base Location is selected as: ${place.name}`;
    }
);

// Updated calculateDistanceAndTime function
function calculateDistanceAndTime(locationName) {
    let origin;
    if (baseLocationInput.value) {
        // Use the base location if it's entered
        origin = baseLocationInput.value;
    } else if (currentLocation) {
        // Use the current location if available
        origin = currentLocation;
    } else {
        alert("Please get your current location or enter a base location first.");
        return;
    }

    const service = new google.maps.DistanceMatrixService();
    service.getDistanceMatrix(
        {
            origins: [origin],
            destinations: [locationName],
            travelMode: "DRIVING",
        },
        function(response, status) {
            if (status === "OK") {
                const distance = response.rows[0].elements[0].distance.text;
                const duration = response.rows[0].elements[0].duration.text;
                const location = { name: locationName, distance: distance, duration: duration };
                locations.push(location);
                updateTable();
            } else {
                alert("Error calculating distance and time: " + status);
            }
        }
    );
}

function removeLocation(index) {
    locations.splice(index, 1);
    updateTable();
}

function updateTable() {
    const tableBody = document.querySelector("#location-table tbody");
    tableBody.innerHTML = "";

    locations.sort((a, b) => {
        const distanceA = parseFloat(a.distance.split(" ")[0]);
        const distanceB = parseFloat(b.distance.split(" ")[0]);
        return distanceA - distanceB;
    });

    locations.forEach((location, index) => {
        const row = tableBody.insertRow();
        const cell1 = row.insertCell(0);
        const cell2 = row.insertCell(1);
        const cell3 = row.insertCell(2);
        const cell4 = row.insertCell(3);
        const cell5 = row.insertCell(4);
        
        cell1.innerHTML = `${index + 1}. ${location.name}`;
        cell2.innerHTML = location.distance;
        cell3.innerHTML = location.duration;
        
        // Create a link to Google Maps for the location
        const googleMapsLink = document.createElement("a");
        googleMapsLink.href = `https://www.google.com/maps?q=${encodeURIComponent(location.name)}`;
        googleMapsLink.target = "_blank"; // Open link in a new tab
        googleMapsLink.textContent = "View on Google Maps";
        cell4.appendChild(googleMapsLink);
      
        // Add a button to remove the location
        const removeButton = document.createElement("button");
        removeButton.textContent = "Remove";
        removeButton.addEventListener("click", () => removeLocation(index));
        cell5.appendChild(removeButton);
    });
}

function downloadExcel() {
    if (locations.length === 0) {
        alert("No data to download.");
        return;
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Locations");

    // Add headers and set alignment for header cells
    const headerRow = worksheet.addRow(["Serial Number", "Location", "Distance", "Travel Time", "Link"]);
    headerRow.eachCell((cell) => {
        cell.alignment = { horizontal: "center" }; // Align headers to center
    });

    // Define cell style for hyperlinks
    const hyperlinkStyle = {
        font: { color: { argb: "0000FF" }, underline: true },
        alignment: { horizontal: "left" }, // Align hyperlinks to the left
    };

    // Add data rows with hyperlinks and set alignment for data cells
    locations.forEach((location, index) => {
        const googleMapsLink = `https://www.google.com/maps?q=${encodeURIComponent(location.name)}`;
        const row = worksheet.addRow([index + 1, location.name, location.distance, location.duration, "Google Maps Link"]);
        const cell = row.getCell(5); // Assuming the "Link" is in the 5th column

        // Apply hyperlink style
        cell.value = {
            text: "Google Maps Link",
            hyperlink: googleMapsLink,
        };
        cell.style = hyperlinkStyle;

        row.eachCell((dataCell) => {
            dataCell.alignment = { horizontal: "center" }; // Align data cells to center
        });
    });

    // Create a Blob containing the Excel file
    workbook.xlsx.writeBuffer().then((buffer) => {
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Create a download link and trigger the download
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'school_locations.xlsx';
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    });
}

// Autocomplete for location input
const locationInput = document.getElementById("location-input");
const autocomplete = new google.maps.places.Autocomplete(locationInput);

autocomplete.addListener("place_changed", function() {
    const place = autocomplete.getPlace();
    if (place && place.name) {
        calculateDistanceAndTime(place.name);
        locationInput.value = "";
    }
});
