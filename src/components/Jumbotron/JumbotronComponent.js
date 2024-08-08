import React, { useState, useEffect } from "react";
import "bootstrap/dist/css/bootstrap.min.css";
import "./JumbotronComponent.css";

const JumbotronComponent = () => {
  const [showZach, setShowZach] = useState(false);

  useEffect(() => {
    // Show the "except for Zach" part after 3 seconds
    const timer = setTimeout(() => {
      setShowZach(true);
    }, 3000);

    // Hide the "except for Zach" part at 12 AM on August 12
    const now = new Date();
    const targetDate = new Date(now.getFullYear(), 7, 12); // August is month 7 (0-based index)
    targetDate.setHours(0, 0, 0, 0); // Set to 12 AM

    const timeUntilTarget = targetDate - now;
    if (timeUntilTarget > 0) {
      const hideTimer = setTimeout(() => {
        setShowZach(false);
      }, timeUntilTarget);

      return () => {
        clearTimeout(hideTimer);
      };
    }

    return () => {
      clearTimeout(timer);
    };
  }, []);

  return (
    <div className="jumbotron jumbotron-fluid jumbotron-custom">
      <div className="container-fluid text-center">
        <h1 className="display-4">Morgo Excel</h1>
        <h4>Making the day faster.{showZach && ".....except for Zach!!"}</h4>
      </div>
    </div>
  );
};

export default JumbotronComponent;
