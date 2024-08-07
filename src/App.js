import React from "react";
import JumbotronComponent from "./components/Jumbotron/JumbotronComponent";
import FileDrop from "./components/FileDrop";
import WeeklyFileDrop from "./components/WeeklyFileDrop";
import "./App.css";
import NavBar from "./components/NavBar/NavBar";
import "bootstrap/dist/css/bootstrap.min.css";

const App = () => {
  return (
    <div className="container-fluid p-0">
      <NavBar /> {/* Changed this line */}
      <JumbotronComponent />
      <div className="row">
        <FileDrop colId="col1" />
        <WeeklyFileDrop />
      </div>
    </div>
  );
};

export default App;
