import { useState } from "react";
import "./App.css";
import axios from "axios";
import { CLIENT_ID, REDIRECT_URL, TENANT_ID } from "./credentials";

function App() {
  const queryParams = new URLSearchParams(window.location.search);
  const [loading, setLoading] = useState(false);
  const [show, setShow] = useState(true);
  const [channel, setChannels] = useState([]);
  const [text, setText] = useState();
  const code = queryParams.get("code");

  const getChannels = async () => {
    setLoading(true);
    setShow(false);
    const url = `https://localhost:5001/teams/?code=${code}`;
    const res = await axios.get(url);
    setChannels(res.data);
    setLoading(false);
  };

  const handleLogin = () => {
    const redirectUrl = REDIRECT_URL;
    const clientId = CLIENT_ID;
    const office365BaseURL = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?response_type=code&client_id=${clientId}&scope=offline_access%20user.read&redirect_uri=${redirectUrl}&state=requestfordevlogin&prompt=login`;
    window.location.href = office365BaseURL;
  };

  return (
    <>
      {loading && (
        <div className="loading">
          <div className="spinner-grow text-dark" role="status">
            <span className="visually-hidden">Loading...</span>
          </div>

          <h3 className="py-3">Getting Channels...</h3>
        </div>
      )}

      {!loading && code && channel && !show && (
        <div className="container p-5">
          <input
            type="text"
            className="form-control my-3"
            placeholder="Search in Teams"
            onChange={(e) => setText(e.target.value)}
          />
          <div className="table-responsive">
            <table className="table">
              <thead>
                <tr>
                  <th scope="col">#</th>
                  <th scope="col">Team Name</th>
                  <th scope="col">Description</th>
                  <th scope="col">Type</th>
                  <th scope="col" className="text-center">
                    Channels
                  </th>
                </tr>
              </thead>
              <tbody>
                {text
                  ? channel
                      .filter((x) => x.name.toLowerCase().includes(text))
                      .map((chnl, index) => (
                        <tr key={index}>
                          <th scope="row">{index + 1}</th>
                          <td>{chnl.name}</td>
                          <td>
                            {chnl.description ? (
                              chnl.description
                            ) : (
                              <p style={{ color: "red" }}>Not Available</p>
                            )}
                          </td>
                          <td>
                            {chnl.visibility ? (
                              chnl.visibility
                            ) : (
                              <p style={{ color: "red" }}>Not Specified</p>
                            )}
                          </td>
                          <td>
                            <ol>
                              {chnl.channels !== null ? (
                                chnl.channels.map((chn, index) => (
                                  <li key={index}>{chn.name}</li>
                                ))
                              ) : (
                                <p style={{ color: "red" }}>Not Available</p>
                              )}
                            </ol>
                          </td>
                        </tr>
                      ))
                  : channel.map((chnl, index) => (
                      <tr key={index}>
                        <th scope="row">{index + 1}</th>
                        <td>{chnl.name}</td>
                        <td>
                          {chnl.description ? (
                            chnl.description
                          ) : (
                            <p style={{ color: "red" }}>Not Available</p>
                          )}
                        </td>
                        <td>
                          {chnl.visibility ? (
                            chnl.visibility
                          ) : (
                            <p style={{ color: "red" }}>Not Specified</p>
                          )}
                        </td>
                        <td>
                          <ol>
                            {chnl.channels !== null ? (
                              chnl.channels.map((chn, index) => (
                                <li key={index}>{chn.name}</li>
                              ))
                            ) : (
                              <p style={{ color: "red" }}>Not Available</p>
                            )}
                          </ol>
                        </td>
                      </tr>
                    ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {code && show && (
        <div className="loading">
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="100"
            height="100"
            fill="currentColor"
            className="bi bi-people-fill"
            viewBox="0 0 16 16"
          >
            <path d="M7 14s-1 0-1-1 1-4 5-4 5 3 5 4-1 1-1 1H7zm4-6a3 3 0 1 0 0-6 3 3 0 0 0 0 6z" />
            <path
              fillRule="evenodd"
              d="M5.216 14A2.238 2.238 0 0 1 5 13c0-1.355.68-2.75 1.936-3.72A6.325 6.325 0 0 0 5 9c-4 0-5 3-5 4s1 1 1 1h4.216z"
            />
            <path d="M4.5 8a2.5 2.5 0 1 0 0-5 2.5 2.5 0 0 0 0 5z" />
          </svg>
          <button className="btn btn-dark my-5" onClick={getChannels}>
            Get Channels
          </button>
        </div>
      )}

      {!code && (
        <div className="loading">
          <svg
            xmlns="http://www.w3.org/2000/svg"
            width="100"
            height="100"
            fill="currentColor"
            className="bi bi-windows"
            viewBox="0 0 16 16"
          >
            <path d="M6.555 1.375 0 2.237v5.45h6.555V1.375zM0 13.795l6.555.933V8.313H0v5.482zm7.278-5.4.026 6.378L16 16V8.395H7.278zM16 0 7.33 1.244v6.414H16V0z" />
          </svg>

          <button className="btn btn-dark my-5" onClick={handleLogin}>
            Sign in with microsoft
          </button>
        </div>
      )}
    </>
  );
}

export default App;
