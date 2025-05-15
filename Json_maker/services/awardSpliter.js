function splitAwards(data) {
    const awards = data.Award.split(/\s+/).filter(Boolean); // Split 'Award' by whitespace and filter out empty strings
    const awardsMap = {}; // Object to hold the result

    awards.forEach((award, index) => {
        awardsMap[`Award${index + 1}`] = award; // Create keys like Award1, Award2, etc.
    });

    return awardsMap;
   
}