import React, { useCallback, useState, useEffect } from "react";
import { useDropzone } from "react-dropzone";
import axios from "axios";
import "bootstrap/dist/css/bootstrap.min.css";

const WeeklyFileDrop = () => {
  const [downloadLink, setDownloadLink] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [showAlert, setShowAlert] = useState(false); // State to control alert visibility

  useEffect(() => {
    const timer = setTimeout(() => {
      setShowAlert(true);
    }, 1500); // 1.5 seconds delay

    return () => clearTimeout(timer); // Cleanup the timer on component unmount
  }, []);

  const onDrop = useCallback((acceptedFiles) => {
    setIsLoading(true);
    setError(null);
    const formData = new FormData();
    formData.append("file", acceptedFiles[0]);

    console.log("Uploading weekly file:", acceptedFiles[0].name);

    axios
      .post("https://api.morgotools.com/upload/weekly", formData, {
        headers: {
          "Content-Type": "multipart/form-data",
        },
      })
      .then((response) => {
        console.log("Response from server:", response.data);
        const filename = response.data.filename;
        const downloadUrl = `https://api.morgotools.com/download/weekly/${filename}`;
        setDownloadLink(downloadUrl);
        console.log("Download link set:", downloadUrl);
      })
      .catch((error) => {
        console.error("Error uploading file:", error);
        setError("Error uploading file. Please try again.");
      })
      .finally(() => {
        setIsLoading(false);
      });
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: ".csv",
  });

  return (
    <div className="col-lg-6">
      <h2 className="text-center">Physical Count File Drop</h2>
      <div className="text-center mt-3">
    <button 
      type="button" 
      className="btn btn-info" 
      style={{ backgroundColor: '#3d6428', color: 'white', borderColor: '#3a4534' }}
      onClick={() => window.open('https://youtu.be/KeFvYRWYhR0', '_blank')}
    >
      Instructional Video
    </button>
  </div>
      <div className="card mt-3" style={{ margin: "30px" }}>
        <div className="card-body">
          <div {...getRootProps({ className: "dropzone" })}>
            <input {...getInputProps()} />
            {isDragActive ? (
              <p>Drop the CSV file here ...</p>
            ) : (
              <p>Drag 'n' drop a CSV file here, or click to select one</p>
            )}
          </div>
          {isLoading && <p>Processing file...</p>}
          {error && <p className="text-danger">{error}</p>}
          {downloadLink && (
            <div>
              <div
                className="alert alert-success"
                role="alert"
                style={{ marginTop: "10px" }}
              >
                <p className="text-center">File processed successfully!</p>
              </div>
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  width: "100%",
                }}
              >
                <a
                  href={downloadLink}
                  download
                  className="btn btn-primary mt-3"
                >
                  Download Processed Weekly File
                </a>
              </div>
            </div>
          )}
          {showAlert && ( // Conditionally render the alert
            <div
              className="alert alert-warning"
              role="alert"
              style={{ marginTop: "30px" }}
            >
              <h4 className="text-center">IMPORTANT</h4>
              <p className="text-center">
                When exporting the Physical Count Sheet, first 'Select None'
                then 'Select All' for the list to render properly.
              </p>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default WeeklyFileDrop;
