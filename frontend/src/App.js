import { useEffect } from "react";
import "@/App.css";

const Home = () => {
  useEffect(() => {
    // The app is a self-contained static page served at /simulator.html.
    // Send the root URL (what the preview opens) straight to it.
    window.location.replace(`${window.location.origin}/simulator.html`);
  }, []);

  return (
    <div className="App">
      <header className="App-header">
        <p className="mt-5">Opening the Factor Investment Simulator…</p>
        <p>
          If it doesn&apos;t load, open{" "}
          <a className="App-link" href="/simulator.html">/simulator.html</a>.
        </p>
      </header>
    </div>
  );
};

function App() {
  return <Home />;
}

export default App;
