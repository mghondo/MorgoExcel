import React from "react";
import { Navbar, Nav, Container } from "react-bootstrap";
import logo from "../../images/VerdantLogo.jpg"; // Adjust the path as needed
import "./Navbar.css"; // Import the CSS file

const NavBar = () => {
  return (
    <Navbar bg="dark" variant="dark" expand="lg">
      {" "}
      {/* Changed bg to "dark" and added variant */}
      <Container>
        <Navbar.Brand
          href="https://www.verdantcreations.com/"
          target="_blank"
          rel="noopener noreferrer"
        >
          <img
            src={logo}
            width="30"
            height="30"
            className="d-inline-block align-top"
            alt="Verdant Logo"
          />
        </Navbar.Brand>
        <Navbar.Toggle aria-controls="basic-navbar-nav" />
        <Navbar.Collapse id="basic-navbar-nav" className="justify-content-end">
          <Nav>
            <Nav.Link
              href="https://pine.backoffice.dutchie.com/"
              target="_blank"
              rel="noopener noreferrer"
            >
              Backoffice
            </Nav.Link>
            <Nav.Link
              href="https://pine.pos.dutchie.com/"
              target="_blank"
              rel="noopener noreferrer"
            >
              Register
            </Nav.Link>
            <Nav.Link
              href="https://www.verdantcreations.com/"
              target="_blank"
              rel="noopener noreferrer"
            >
              Shop Verdant
            </Nav.Link>
          </Nav>
        </Navbar.Collapse>
      </Container>
    </Navbar>
  );
};

export default NavBar;
