// JavaScript
let currentLocation;
let locations = [];

function getCurrentLocation() {
    if (navigator.geolocation) {
        navigator.geolocation.getCurrentPosition(function(position) {
            currentLocation = new google.maps.LatLng(position.coords.latitude, position.coords.longitude);
            document.getElementById("current-location").textContent = `Obtained your Current Location: ${currentLocation}`;
        });
    } else {
        alert("Geolocation is not supported by this browser.");
    }
}

function calculateDistanceAndTime(locationName) {
    if (!currentLocation) {
        alert("Please get your current location first.");
        return;
    }

    const service = new google.maps.DistanceMatrixService();
    service.getDistanceMatrix(
        {
            origins: [currentLocation],
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

    const data = [["Serial Number","Location", "Distance", "Travel Time", "Link"]];
    locations.forEach((location,index) => {
        const googleMapsLink = `https://www.google.com/maps?q=${encodeURIComponent(location.name)}`;
        // Create a hyperlink formula for the "Link" column
        const hyperlinkFormula = `=HYPERLINK("${googleMapsLink}", "Google Maps Link")`;
        data.push([index+1, location.name, location.distance, location.duration, hyperlinkFormula]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Locations");

    // Generate Excel file and trigger download
    XLSX.writeFile(wb, "school_locations.xlsx");
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
